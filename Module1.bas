Attribute VB_Name = "Module1"
#If Win64 Then
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function GetCursorPos Lib "user32" (p As theCursor) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As LongPtr
#Else
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (p As theCursor) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Lonr, ByVal hdc As Long) As Long
#End If

Const LOGPIXELSX = 88
Const LOGPIXELSY = 90

Public Type theCursor

    Left As Long
    Top As Long
    
End Type

Sub MoveFormToMouse()

    Dim mousePos As theCursor
    
    GetCursorPos mousePos
    UserForm1.Left = pointsPerPixelY * mousePos.Left
    UserForm1.Top = pointsPerPixelX * mousePos.Top
    
End Sub

Private Function pointsPerPixelX() As Double

    hdc = GetDC(0)
    pointsPerPixelX = 72 / GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc
    
End Function

Private Function pointsPerPixelY() As Double

    hdc = GetDC(0)
    pointsPerPixelY = 72 / GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
    
End Function


