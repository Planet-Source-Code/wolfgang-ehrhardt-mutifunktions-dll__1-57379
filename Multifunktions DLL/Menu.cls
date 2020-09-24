VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCountA Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, Mi As MENUINFO) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type MENUITEMINFO
    cbSize        As Long
    fMask         As Long
    fType         As Long
    fState        As Long
    wID           As Long
    hSubMenu      As Long
    hbmpChecked   As Long
    hbmpUnchecked As Long
    dwItemData    As Long
    dwTypeData    As String
    cch           As Long
End Type

Private Type MENUINFO
    cbSize          As Long
    fMask           As Long
    dwStyle         As Long
    cyMax           As Long
    hbrBack         As Long
    dwContextHelpID As Long
    dwMenuData      As Long
End Type

Public Enum MenuNFO
    mMENUBAR = 1
    mMENU = 2
    mSYSMENU = 3
End Enum

Private Const MIM_BACKGROUND As Long = &H2
Private Const MIM_APPLYTOSUBMENUS As Long = &H80000000

Dim FoundMenuItem As Boolean
Public Function Get_MenuHwnd(ByVal hWnd As Long) As Long
    Get_MenuHwnd = GetMenu(hWnd)
End Function
Public Function Get_SystemMenuHwnd(ByVal hWnd As Long) As Long
    Get_SystemMenuHwnd = GetSystemMenu(hWnd, False)
End Function
Public Function Get_MenuItemCount(ByVal hWnd As Long) As Long
    Get_MenuItemCount = GetMenuItemCountA(Get_MenuHwnd(hWnd))
End Function
Public Sub Find_MenuItemByCaption(ByVal hWnd As Long, _
                                  ByVal MenuItemText As String, _
                                  mID2Store As Long, _
                                  mHwnd2Store As Long, _
                                  Pos2Store As Long, _
                                  Optional Systemmenu As Boolean = False)
    Dim mHwnd As Long
    
    FoundMenuItem = False
    
    mHwnd = IIf(Systemmenu, _
                GetSystemMenu(hWnd, False), _
                GetMenu(hWnd))
    
    Call SMenu(mHwnd, mID2Store, mHwnd2Store, Pos2Store, _
               False, MenuItemText, hWnd)
    
End Sub
Public Function Click_MenuItemByCaption(ByVal hWnd As Long, _
                                        ByVal Caption As String) _
                                                    As Boolean
    Dim Mid As Long, mHwnd As Long, Pos As Long
    
    Call Find_MenuItemByCaption(hWnd, Caption, Mid, mHwnd, Pos)
    Click_MenuItemByCaption = Click_MenuItemByID(hWnd, Mid)

End Function
Public Function Click_MenuItemByID(ByVal hWnd As Long, _
                                   ByVal Mid As Long) As Boolean
    Click_MenuItemByID = PostMessage(hWnd, WM_COMMAND, Mid, 0&)
End Function
Public Sub Click_MenuItemBymHwnd(ByVal hWnd As Long, _
                                 ByVal mHwnd As Long, _
                                 ByVal Pos As Long)
    Call Click_MenuItemByID(hWnd, GetMenuItemID(mHwnd, Pos))
End Sub
Private Function SMenu(ByVal mHwnd As Long, _
                       mID2Store As Long, _
                       mHwnd2Store As Long, _
                       Pos As Long, _
                       Optional iSubFlag As Boolean, _
                       Optional CaptionToFind As String, _
                       Optional Owner As Long) As Long
    Dim n As Long, I As Long, id As Long
    Dim Buffer As String
    Dim Result As Long
    
    On Local Error Resume Next
    
    If FoundMenuItem Then Exit Function
    
    n = GetMenuItemCountA(mHwnd)
    
    For I = 0 To n - 1
        id = GetMenuItemID(mHwnd, I)
        
        If id > -1 Then
            Buffer = Space$(128)
            Result = GetMenuString(mHwnd, id, Buffer, _
                                   Len(Buffer), MF_BYCOMMAND)
            Buffer = Left$(Buffer, Result)
            
            If Buffer = CaptionToFind Then _
                mID2Store = id: _
                mHwnd2Store = mHwnd: _
                Pos = I: _
                FoundMenuItem = True: _
                Exit Function
        End If
    
        If Not iSubFlag Then
            If GetSubMenu(mHwnd, I) > 1 Then _
                iSubFlag = True: _
                SMenu GetSubMenu(mHwnd, I), mID2Store, mHwnd2Store, Pos, True, CaptionToFind, Owner: _
                iSubFlag = False
        Else
            If GetSubMenu(mHwnd, I) > 1 Then _
                iSubFlag = True: _
                SMenu GetSubMenu(mHwnd, I), mID2Store, mHwnd2Store, Pos, True, CaptionToFind, Owner
        End If
        
        If FoundMenuItem Then Exit Function
   Next I
    
End Function
Public Function Remove_MenuByCaption(ByVal hWnd As Long, _
                                     ByVal Caption As String, _
                                     Optional Systemmenu As Boolean = False) _
                                             As Boolean
    Dim Mid As Long, mHwnd As Long, Pos As Long
    
    Call Find_MenuItemByCaption(hWnd, Caption, _
                                Mid, mHwnd, Pos, Systemmenu)
    
    If RemoveMenuA(mHwnd, Pos, MF_BYPOSITION) Then _
        Remove_MenuByCaption = True

End Function
Public Function Set_MenuColor(SetWhat As MenuNFO, _
                              ByVal hWnd As Long, _
                              ByVal Color As Long, _
                              Optional MenuIndex As Integer, _
                              Optional IncludeSubmenus As Boolean = False) _
                                    As Boolean
    
    Dim Mi As MENUINFO
    Dim clrref As Long, hSysMenu As Long, mHwnd As Long
         
    On Local Error GoTo Quit
   
    clrref = GRAFX.Convert_OLEtoRBG(Color)
   
    Mi.cbSize = Len(Mi)
    Mi.hbrBack = CreateSolidBrush(clrref)
    
    Select Case SetWhat
        Case mMENUBAR
            Mi.fMask = MIM_BACKGROUND
            Call SetMenuInfo(GetMenu(hWnd), Mi)
        Case mMENU
            If MenuIndex = 0 Then _
                Set_MenuColor = Set_MenuColor(mMENUBAR, hWnd, Color): _
                Exit Function
            
            If MenuIndex < 1 _
            Or Get_MenuItemCount(hWnd) < MenuIndex Then _
                Exit Function
    
            Mi.fMask = IIf(IncludeSubmenus, _
                           MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS, _
                           MIM_BACKGROUND)
            
            mHwnd = GetMenu(hWnd)
            mHwnd = GetSubMenu(mHwnd, MenuIndex - 1)
        
            Call SetMenuInfo(mHwnd, Mi)
            hWnd = mHwnd
        Case mSYSMENU
            hSysMenu = GetSystemMenu(hWnd, False)
   
            Mi.fMask = MIM_BACKGROUND _
                       Or MIM_APPLYTOSUBMENUS
            
            Call SetMenuInfo(hSysMenu, Mi)
            hWnd = hSysMenu
    End Select
    
    Call DrawMenuBar(hWnd)
    Set_MenuColor = True
    
Quit:

End Function
Public Function AddPictureToMenu(ByVal hWnd As Long, _
                                 ByVal MenuText As String, _
                                 Picture As StdPicture, _
                                 Optional Systemmenu As Boolean = False) _
                                      As Boolean
        Dim R As Long, S As Long, T As Long
        
        On Local Error GoTo Quit
        
        Call Find_MenuItemByCaption(hWnd, _
                                    MenuText, R, S, T, _
                                    Systemmenu)
        
        If S Then _
            If SetMenuItemBitmaps(S, T, MF_BYPOSITION, _
                                  Picture, Picture) Then _
                                        AddPictureToMenu = True
Quit:
End Function
