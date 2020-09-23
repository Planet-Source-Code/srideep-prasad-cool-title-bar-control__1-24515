Attribute VB_Name = "Module2"
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SIZE_MAXIMIZED = 2
Public Const SC_MAXIMIZE = &HF030&


Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type



Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOSENDCHANGING = &H400

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const WM_KILLFOCUS = &H8
Public Const WM_SETFOCUS = &H7
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_STYLECHANGED = &H7D
Public Const WM_SETICON = &H80
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_MOVING = &H216
Public Const WM_ACTIVATE = &H6
Public Const WM_ACTIVATEAPP = &H1C
Public Const WA_ACTIVE = 1
Public Const WA_INACTIVE = 0
Public Const WM_PAINT = &HF
Public Const WA_CLICKACTIVE = 2
Public Const WM_NCPAINT = &H85
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WS_VISIBLE = &H10000000
Public Const WM_NCACTIVATE = &H86
Public Const WM_CLOSE = &H10
Public Const SC_CLOSE = &HF060&
Public Const WM_SYSCOMMAND = &H112
Public Const WM_APP = &H8000
Public Const SM_CXMAXIMIZED = 61
Public Const SM_CYMAXIMIZED = 62
Public Const WM_SHOWWINDOW = &H18
Public Const HWND_BOTTOM = 1
Public Const WM_MDIACTIVATE = &H222





Sub SetTopMost(mhWnd As Long, ByVal Value As Boolean)
   Const swpFlags = SWP_NOMOVE Or SWP_NOSIZE
   If Value Then
      Call SetWindowPos(mhWnd, HWND_TOPMOST, 0, 0, 0, 0, swpFlags)
   Else
      Call SetWindowPos(mhWnd, HWND_NOTOPMOST, 0, 0, 0, 0, swpFlags)
   End If
   
End Sub





Function LoWord(ByRef lThis As Long) As Long
   LoWord = (lThis And &HFFFF&)
End Function

Function HiWord(ByRef lThis As Long) As Long
   If (lThis And &H80000000) = &H80000000 Then
      HiWord = ((lThis And &H7FFF0000) \ &H10000) Or &H8000&
   Else
      HiWord = (lThis And &HFFFF0000) \ &H10000
   End If
End Function



Function TitleBarProc(ByVal hwnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
    TitleBarProc = CallWindowProc(TitleBarProcOld, hwnd, uMsg, wParam, lParam)
    
    If uMsg = WM_MOVE Then
        If ParentState <> vbMinimized Then
            TitleBarForm.RefreshOnMove
        End If
    End If
    
End Function


Function GetTopMost(mhWnd As Long) As Boolean
   ' Return value of WS_EX_TOPMOST bit.
   GetTopMost = CBool(fStyleEx(mhWnd) And WS_EX_TOPMOST)
End Function
Private Function fStyleEx(mhWnd As Long, Optional ByVal NewBits As Long = 0) As Long
   '
   ' Set new extended style bits.
   '
   If NewBits Then
      Call SetWindowLong(mhWnd, GWL_EXSTYLE, NewBits)
   End If
   ' Retrieve current extended style bits.
   fStyleEx = GetWindowLong(mhWnd, GWL_EXSTYLE)
End Function

