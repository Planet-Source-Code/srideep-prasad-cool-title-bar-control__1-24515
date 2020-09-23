Attribute VB_Name = "Module1"
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function GetRValue Lib "GDI32" (ByVal Color As Long) As Integer
Public Declare Function GetGValue Lib "GDI32" (ByVal Color As Long) As Integer
Public Declare Function GetBValue Lib "GDI32" (ByVal Color As Long) As Integer




Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    Dim CLR_INVALID As Long
    CLR_INVALID = -1
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


Function UnRGB(ColorVal As Long, Part As Long) As Long
Dim Num As Integer, I As Integer
H = Trim$(Hex$(ColorVal))
For I = 1 To (6 - Len(H))
    H = "0" + H
Next I
Select Case 2 - Part
    Case Is = 0
        hexnum$ = Mid$(H, 1, 2)
    Case Is = 1
        hexnum$ = Mid$(H, 3, 2)
    Case Is = 2
        hexnum$ = Mid$(H, 5, 2)
End Select
outhex$ = UCase$(hexnum$)         ' Convert characters to uppercase
Num = Val("&h" + outhex$)      ' &H tells VAL that A,B,C,to F are OK
UnRGB = Num

End Function
