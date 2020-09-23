VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Win32 APIs used to toggle border styles.
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Win32 APIs used to automate drag and sysmenu support.
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
    
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" _
    (lpPoint As POINTAPI) As Long

Private Type POINTAPI
   x As Long
   y As Long
End Type

' Used to support captionless drag
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

' Undocumented message constant.
Private Const WM_GETSYSMENU = &H313

' Used to get window style bits.
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

' Style bits.
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000
Private Const WS_SYSMENU = &H80000
Private Const WS_CAPTION = &HC00000
Private Const DS_MODALFRAME = &H80         '  Can be combined with WS_CAPTION

' Extended Style bits.
Private Const WS_EX_TOPMOST = &H8
Private Const WS_EX_TOOLWINDOW = &H80
Private Const WS_EX_CONTEXTHELP = &H400
Private Const WS_VISIBLE = &H10000000

' Force total pRedraw that shows new styles.
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOREDRAW = &H8

' Used to toggle into topmost layer.
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

' References to client form.
Private WithEvents mClient As Form
Attribute mClient.VB_VarHelpID = -1
Private mhWnd As Long
Dim ParentForm As Form
' Class level variables
Private mAutoSysMenu As Boolean
Private mAutoDrag As Boolean

Private Sub mClient_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   '
   ' Let user drag form using any portion of the form,
   ' not just titlebar, when user depresses left mousebutton.
   ' Useful for captionless forms.
   '
   If Button = vbLeftButton Then
      If mAutoDrag Then
         Call ReleaseCapture
         Call SendMessage(mhWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
         
      End If
   End If
30 End Sub

Private Sub mClient_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   '
   ' Handle system menu display when user right-clicks
   ' anywhere on form. Useful for captionless forms.
   '
   Dim PT As POINTAPI
   '
   ' This is relative to the screen, so we can't
   ' use the coordinates passed in the event
   '
   Call GetCursorPos(PT)
   If Button = vbRightButton Then
      If mAutoSysMenu Then
         Call ShowSysMenu(PT.x, PT.y)
      End If
   End If
End Sub
Sub SetAutoDrag(ByVal Value As Boolean)
    '
    ' Let the user move the form using any portion of form,
    ' not just titlebar, when user depresses left mousebutton.
    ' Useful for captionless forms.
    mAutoDrag = Value
End Sub
Function GetAutoDrag() As Boolean
   
   GetAutoDrag = mAutoDrag

End Function
Sub SetAutoSysMenu(ByVal Value As Boolean)
    '
    ' Handle system menu display when user right-clicks
    ' anywhere on form. Useful for captionless forms.
    '
    mAutoSysMenu = Value
End Sub

Function GetAutoSysMenu() As Boolean
   GetAutoSysMenu = mAutoSysMenu
End Function

Public Sub SetClient(ByVal obj As Form)
   ' Store reference to client form.
   Set mClient = obj
   ' Cache hWnd as it'll be accessed frequently.
   If mClient Is Nothing Then
      mhWnd = 0
   Else
      mhWnd = mClient.hwnd
   End If
End Sub

Function GetClient() As Form
   ' Return reference to client form.
   Set GetClient = mClient
End Function

Public Property Let ControlBox(ByVal Value As Boolean)
   ' Set WS_SYSMENU On or Off as requested.
   Call fFlipBit(WS_SYSMENU, Value)
End Property

Public Property Get ControlBox() As Boolean
   ' Return value of WS_SYSMENU bit.
   ControlBox = CBool(fStyle And WS_SYSMENU)
End Property

Public Property Let MaxButton(ByVal Value As Boolean)
   ' Set WS_MAXIMIZEBOX On or Off as requested.
   Call fFlipBit(WS_MAXIMIZEBOX, Value)
End Property

Public Property Get MaxButton() As Boolean
   ' Return value of WS_MAXIMIZEBOX bit.
   MaxButton = CBool(fStyle And WS_MAXIMIZEBOX)
End Property

Public Property Let MinButton(ByVal Value As Boolean)
   ' Set WS_MINIMIZEBOX On or Off as requested.
   Call fFlipBit(WS_MINIMIZEBOX, Value)
End Property

Public Property Get MinButton() As Boolean
   ' Return value of WS_MINIMIZEBOX bit.
   MinButton = CBool(fStyle And WS_MINIMIZEBOX)
End Property

Public Property Let Sizable(ByVal Value As Boolean)
   ' Set WS_THICKFRAME On or Off as requested.
   Call fFlipBit(WS_THICKFRAME, Value)
End Property

Public Property Get Sizable() As Boolean
   ' Return value of WS_THICKFRAME bit.
   Sizable = CBool(fStyle And WS_THICKFRAME)
End Property

Public Property Let Titlebar(ByVal Value As Boolean)
   ' Set WS_CAPTION On or Off as requested.
   Call fFlipBit(WS_CAPTION, Value)
End Property

Public Property Get Titlebar() As Boolean
   ' Return value of WS_CAPTION bit.
   Titlebar = CBool(fStyle And WS_CAPTION)
End Property



Public Property Let ToolWindow(ByVal Value As Boolean)
   ' Set WS_EX_TOOLWINDOW On or Off as requested.
   Call fFlipBitEx(WS_EX_TOOLWINDOW, Value)
End Property

Public Property Get ToolWindow() As Boolean
   ' Return value of WS_EX_TOOLWINDOW bit.
   ToolWindow = CBool(fStyleEx And WS_EX_TOOLWINDOW)
End Property

Sub SetTopMost(ByVal Value As Boolean)
   Const swpFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
   '
   ' Unlike most style bits, WS_EX_TOPMOST must be
   ' set with SetWindowPos rather than SetWindowLong.
   '
   If Value Then
      Call SetWindowPos(mhWnd, HWND_TOPMOST, 0, 0, 0, 0, swpFlags)
   Else
      Call SetWindowPos(mhWnd, HWND_NOTOPMOST, 0, 0, 0, 0, swpFlags)
   End If
   ' Additional references on VB use of SetWindowPos...
   ' BUG: SetWindowPos API Does Not Set Topmost Window in VB
   ' -- http://support.microsoft.com/support/kb/articles/Q192/2/54.ASP
   ' FIX: TopMost Window Does Not Stay on Top in Design Environment
   ' -- http://support.microsoft.com/support/kb/articles/Q150/2/33.ASP
End Sub

Function GetTopMost() As Boolean
   ' Return value of WS_EX_TOPMOST bit.
   GetTopMost = CBool(fStyleEx And WS_EX_TOPMOST)
End Function

Public Property Let WhatsThisButton(ByVal Value As Boolean)
   ' Set WS_EX_CONTEXTHELP On or Off as requested.
   Call fFlipBitEx(WS_EX_CONTEXTHELP, Value)
End Property

Public Property Get WhatsThisButton() As Boolean
   ' Return value of WS_EX_CONTEXTHELP bit.
   WhatsThisButton = CBool(fStyleEx And WS_EX_CONTEXTHELP)
End Property

Public Property Get hwnd() As Long
   hwnd = mhWnd
End Property

Public Sub pRedraw()
   ' Redraw window with new style.
   Const swpFlags As Long = _
      SWP_FRAMECHANGED Or SWP_NOMOVE Or _
      SWP_NOZORDER Or SWP_NOSIZE
   Call SetWindowPos(mhWnd, 0, 0, 0, 0, 0, swpFlags)
End Sub

Public Sub ShowSysMenu(ByVal x As Long, ByVal y As Long)
   ' Must be in screen coordinates.
   Call SendMessage(mhWnd, WM_GETSYSMENU, 0, ByVal fMakeLong(y, x))
End Sub

Private Function fMakeLong(ByVal WordHi As Integer, ByVal WordLo As Integer) As Long
    '
    ' High word is coerced to a Long to allow it to
    ' overflow limits of multiplication which shifts
    ' it left.
    '
    fMakeLong = (CLng(WordHi) * &H10000) Or (WordLo And &HFFFF&)
End Function

Private Function fStyle(Optional ByVal NewBits As Long = 0) As Long
   '
   ' Set new style bits.
   '
   If NewBits Then
      Call SetWindowLong(mhWnd, GWL_STYLE, NewBits)
   End If
   ' Retrieve current style bits.
   fStyle = GetWindowLong(mhWnd, GWL_STYLE)
End Function

Private Function fStyleEx(Optional ByVal NewBits As Long = 0) As Long
   '
   ' Set new extended style bits.
   '
   If NewBits Then
      Call SetWindowLong(mhWnd, GWL_EXSTYLE, NewBits)
   End If
   ' Retrieve current extended style bits.
   fStyleEx = GetWindowLong(mhWnd, GWL_EXSTYLE)
End Function

Private Function fFlipBit(ByVal Bit As Long, ByVal Value As Boolean) As Boolean
Dim lStyle As Long
   
   ' Retrieve current style bits.
   lStyle = GetWindowLong(mhWnd, GWL_STYLE)
   
   ' Set requested bit On or Off and Redraw.
   If Value Then
      lStyle = lStyle Or Bit
   Else
      lStyle = lStyle And Not Bit
   End If
   Call SetWindowLong(mhWnd, GWL_STYLE, lStyle)
   Call pRedraw
   
   ' Return success code.
   fFlipBit = (lStyle = GetWindowLong(mhWnd, GWL_STYLE))
End Function

Private Function fFlipBitEx(ByVal Bit As Long, ByVal Value As Boolean) As Boolean
Dim lStyleEx As Long
   
   ' Retrieve current extended style bits.
   lStyleEx = GetWindowLong(mhWnd, GWL_EXSTYLE)
   
   ' Set requested bit On or Off and Redraw.
   If Value Then
      lStyleEx = lStyleEx Or Bit
   Else
      lStyleEx = lStyleEx And Not Bit
   End If
   Call SetWindowLong(mhWnd, GWL_EXSTYLE, lStyleEx)
   Call pRedraw
   
   ' Return success code.
   fFlipBitEx = (lStyleEx = GetWindowLong(mhWnd, GWL_EXSTYLE))
End Function

Friend Sub MoveParent(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
         Call ReleaseCapture
         Call SendMessage(mhWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
   End If
End Sub

Public Function IsModal() As Boolean
    IsModal = CBool(fStyle And DS_MODALFRAME)
End Function


Sub SetParent(ByVal P As Form)
Set ParentForm = P
End Sub

