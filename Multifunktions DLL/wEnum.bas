Attribute VB_Name = "wEnum"
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long

Public tTask() As Long, tChild() As Long
Private Function WndEnumProc(ByVal Hwnd As Long, ByVal lParam As Long) As Long
        
    ReDim Preserve tTask(UBound(tTask) + 1)
    tTask(UBound(tTask)) = Hwnd
        
    WndEnumProc = 1
            
End Function
Private Function WndEnumChildProc(ByVal Hwnd As Long, ByVal lParam As Long) As Long
    ReDim Preserve tChild(UBound(tChild) + 1)
    tChild(UBound(tChild)) = Hwnd
    WndEnumChildProc = 1
End Function
Public Sub EnumTask()
        
    ReDim tTask(0)
    Call EnumWindows(AddressOf WndEnumProc, CLng(1))
    
    Call ARR.DeleteIndex(tTask, 0)
    
End Sub
Public Sub EnumChild(ByVal Hwnd As Long)
        
    ReDim tChild(0)
    Call EnumChildWindows(ByVal Hwnd, AddressOf WndEnumChildProc, CLng(0))

    Call ARR.DeleteIndex(tChild, 0)
    
End Sub
Public Sub EnumAll(ARR As Variant)
    
    On Local Error Resume Next
    
    Call EnumTask
    
    ARR = tTask

End Sub
