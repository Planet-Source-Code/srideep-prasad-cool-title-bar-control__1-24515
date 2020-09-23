VERSION 5.00
Begin VB.Form CapFrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   30
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   2
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin CoolCaptionControl.GradientFX Ing 
      Left            =   4260
      Top             =   -90
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin CoolCaptionControl.GradientFX GF 
      Left            =   4740
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      GradColor1      =   16711680
      GradColor2      =   16761024
   End
   Begin CoolCaptionControl.CoolCommand CloseBtn 
      Height          =   270
      Left            =   6585
      TabIndex        =   3
      Top             =   30
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      checkcaption    =   "r"
      Caption         =   "r"
   End
   Begin CoolCaptionControl.CoolCommand MaxRest 
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   30
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      checkcaption    =   "1"
      Caption         =   "1"
   End
   Begin CoolCaptionControl.CoolCommand Min 
      Height          =   270
      Left            =   5880
      TabIndex        =   1
      Top             =   30
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      checkcaption    =   "0"
      Caption         =   "0"
   End
   Begin VB.Image FrmIcon 
      Height          =   315
      Left            =   15
      Picture         =   "frmCap.frx":0000
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   300
   End
   Begin VB.Label Cap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Captions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   330
      TabIndex        =   0
      Top             =   45
      Width           =   750
   End
   Begin VB.Image Img 
      Height          =   300
      Left            =   30
      Stretch         =   -1  'True
      Top             =   30
      Width           =   7095
   End
End
Attribute VB_Name = "CapFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const GWL_WNDPROC = (-4)


Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Const SM_CYCAPTION = 4
Const SM_CYSMCAPTION = 51
Const SM_CYBORDER = 6
Const SM_CYDLGFRAME = 8
Const SM_CYFRAME = 33
Const SM_CYMENU = 15

Dim GStyle As Long

Dim Rthreshold As Integer

Dim NCapSize As Long
Dim SCapSize As Long
Dim sBorder As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Dim BARHEIGHT As Long
Dim CAPFLAG As Boolean
Dim SCALEUNIT As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Dim MFlag As Boolean
Dim STimer As Long
Dim State As Long
Const CloseChar = "r"
Const MinBtn = "0"
Const Max = "1"
Const Rest = "2"
Dim WC As clsStyle
Dim Title As New clsStyle
Dim WithEvents Pfrm As Form
Attribute Pfrm.VB_VarHelpID = -1

Dim TitleBarForm As Form
Dim ParentState As Long
Dim ParentFrm As Form
Dim MoveParent As Boolean
Dim mFix As Boolean
Dim pRect As RECT
Dim mnuHeight As Long
Dim Active As Boolean
Dim IC1 As OLE_COLOR, IC2 As OLE_COLOR
Dim InacGrad As Boolean
Dim AcForeColor As OLE_COLOR, InTxt As OLE_COLOR
Dim ModalFlag As Boolean
Dim ByPassRefresh As Boolean, Init As Boolean
Dim pLeft As Long, pTop As Long, MaxDim As Long
Dim dTitle As RECT, pVisible As Boolean, aFlag As Boolean
Dim aPic As IPictureDisp, inacPic As IPictureDisp
Implements ISubclass

Sub SetParent(Frm As Form)
    Set Pfrm = Frm
    Set ParentFrm = Frm
    Set WC = New clsStyle
    Set ParentForm = Frm
    PHwnd = Pfrm.hwnd
    
    WC.SetParent Frm
    
    WC.SetClient Pfrm
    If WC.Titlebar = False Then
        Set Pfrm = Nothing
        Set WC = Nothing
        Unload Me
    End If
    Call GetBorders
    Init = False
    GetWindowRect Pfrm.hwnd, pRect
    If WC.ToolWindow = False Then
        SetWindowPos Pfrm.hwnd, 0&, pRect.Left, pRect.Top, pRect.Right - pRect.Left, pRect.Bottom - pRect.Top - sBorder - NCapSize, SWP_NOACTIVATE Or SWP_NOZORDER
    Else
        SetWindowPos Pfrm.hwnd, 0&, pRect.Left, pRect.Top, pRect.Right - pRect.Left, pRect.Bottom - pRect.Top - sBorder - SCapSize, SWP_NOACTIVATE Or SWP_NOZORDER
    End If
    Init = False
    BARHEIGHT = Pfrm.Height - Pfrm.ScaleHeight
    CAPFLAG = WC.Titlebar
    Img.Enabled = False
    Title.SetClient Me
'   Title.SetTopMost True
    Title.SetAutoDrag True
    AttachMessage Me, Me.hwnd, WM_MOVE
    AttachMessage Me, Me.hwnd, WM_ACTIVATE
    AttachMessage Me, Pfrm.hwnd, WM_STYLECHANGED
    AttachMessage Me, Pfrm.hwnd, WM_SETICON
    AttachMessage Me, Pfrm.hwnd, WM_ACTIVATE
    AttachMessage Me, Pfrm.hwnd, WM_MOVE
    AttachMessage Me, Pfrm.hwnd, WM_APP + 1
    AttachMessage Me, Pfrm.hwnd, WM_PAINT
    AttachMessage Me, Pfrm.hwnd, WM_WINDOWPOSCHANGED
    
            
End Sub






Private Sub Cap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Pfrm.WindowState = vbMaximized Then GoTo 10
'Timer1.Enabled = True
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)

10 End Sub

Private Sub CloseBtn_Click()
    SendMessage Pfrm.hwnd, WM_APP + 1, 0, 0
End Sub

Private Sub Form_Click()
BringWindowToTop Pfrm.hwnd
End Sub

Private Sub Form_DblClick()
'    Title.SetTopMost True
If MaxRest.Enabled = True Then Call MaxRest_Click
End Sub


Private Sub Form_Load()
    Me.Visible = False
    Set TitleBarForm = Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Timer1.Enabled = True
If Pfrm.WindowState = vbMaximized Then
    GoTo 20
End If

20 End Sub


Private Sub Form_Resize()
Cap.Top = (Me.ScaleHeight - Cap.Height) / 2
FrmIcon.Top = (Me.ScaleHeight - FrmIcon.Height) / 2
Img.Height = Me.ScaleHeight
End Sub

Sub RefreshAll()
    If ByPassRefresh = True Then Exit Sub
    ByPassRefresh = True
    MaxDim = GetSystemMetrics(SM_CYMAXIMIZED)
    If WC.Titlebar = True Then WC.Titlebar = False
     If IsWindowEnabled(Me.hwnd) = 0 Then
        EnableWindow Me.hwnd, True
        ModalFlag = True
     End If
    Call GetBorders
    Call CheckWinBtns
    Call GetBorders
    
    Call GetWindowRect(Pfrm.hwnd, pRect)
    
'   If Pfrm.Visible = False Then
'       If Me.Visible = True Then Me.Visible = False
'   Else
'       If Me.Visible = False Then Me.Visible = True
'   End If
    
    If Pfrm.Visible Then
        If Me.Visible = False Then Me.Visible = True
    End If
       
    If WC.ToolWindow = False Then
'       Me.Height = ((NCapSize + sBorder) * Screen.TwipsPerPixelY)
        SetWindowPos Me.hwnd, 0&, pRect.Left, pRect.Top - (NCapSize + sBorder), pRect.Right - pRect.Left, (NCapSize + sBorder), SWP_NOACTIVATE
    Else
'       Me.Height = (SCapSize + sBorder) * Screen.TwipsPerPixelY
        SetWindowPos Me.hwnd, 0&, pRect.Left, pRect.Top - (SCapSize + sBorder), pRect.Right - pRect.Left, (SCapSize + sBorder), SWP_NOACTIVATE
    End If
      
   
    If pRect.Bottom - pRect.Top > MaxDim Then
        Dim Diff As Long
        Diff = (pRect.Bottom - pRect.Top) - MaxDim
        If WC.ToolWindow = False Then
            SetWindowPos Pfrm.hwnd, 0&, pRect.Left, (NCapSize + sBorder), pRect.Right - pRect.Left, (pRect.Bottom - pRect.Top) - (sBorder * 2) - NCapSize - Diff, SWP_NOACTIVATE
        Else
            SetWindowPos Pfrm.hwnd, 0&, pRect.Left, (SCapSize + sBorder), pRect.Right - pRect.Left, (pRect.Bottom - pRect.Top) - (sBorder * 2) - SCapSize - Diff, SWP_NOACTIVATE
        End If
    End If
    
        
    Img.Left = 0
    FrmIcon.Left = 0
    CloseBtn.Left = Me.ScaleWidth - CloseBtn.Width
    MaxRest.Left = Me.ScaleWidth - CloseBtn.Width - MaxRest.Width
    Min.Left = Me.ScaleWidth - CloseBtn.Width - MaxRest.Width - Min.Width

    Cap.Top = (Me.ScaleHeight - Cap.Height) / 2
    FrmIcon.Top = (Me.ScaleHeight - FrmIcon.Height) / 2
    
    
    
    
    
    Img.Width = Me.ScaleWidth
    Img.Height = Me.ScaleHeight
    
    Min.Height = CloseBtn.Height
    MaxRest.Height = CloseBtn.Height
    
    CloseBtn.Top = (Me.ScaleHeight - CloseBtn.Height) / 2
    MaxRest.Top = (Me.ScaleHeight - MaxRest.Height) / 2
    Min.Top = (Me.ScaleHeight - Min.Height) / 2
    
    Set FrmIcon.Picture = Pfrm.Icon
    
    Select Case Pfrm.WindowState
        Case Is = vbMinimized
            Me.Hide
        Case Is = vbMaximized
            Title.SetAutoDrag False
            MaxRest.Caption = "2"
        Case Is = vbNormal
            MaxRest.Caption = "1"
            Title.SetAutoDrag True
    End Select
    If Init = True Then
    If Active = True Then
        InacGrad = False
        Dim GC1 As OLE_COLOR, GC2 As OLE_COLOR
            GC1 = GF.GradColor1
            GC2 = GF.GradColor2
            Me.Cls
            
                Select Case GStyle
                    Case Is = 1
                        GF.DoGradientHorizontal Me
                    Case Is = 2
                        GF.DoGradientVertical Me
                End Select
    Else
        Select Case GStyle
            Case Is = 1
                Ing.DoGradientHorizontal Me
            Case Is = 2
                Ing.DoGradientVertical Me
        End Select
    End If
    End If
    ByPassRefresh = False
End Sub



Private Sub Form_Unload(Cancel As Integer)
    DetachMessage Me, Me.hwnd, WM_MOVE
    DetachMessage Me, Me.hwnd, WM_ACTIVATE
    DetachMessage Me, Pfrm.hwnd, WM_STYLECHANGED
    DetachMessage Me, Pfrm.hwnd, WM_SETICON
    DetachMessage Me, Pfrm.hwnd, WM_ACTIVATE
    DetachMessage Me, Pfrm.hwnd, WM_MOVE
    DetachMessage Me, Pfrm.hwnd, WM_APP + 1
    DetachMessage Me, Pfrm.hwnd, WM_PAINT
    DetachMessage Me, Pfrm.hwnd, WM_WINDOWPOSCHANGED
    Init = False
    ByPassRefresh = True
    WC.Titlebar = True
    Pfrm.Caption = Me.Cap.Caption
    Set Pfrm.Icon = Me.FrmIcon.Picture
    
End Sub

Private Sub FrmIcon_Click()
    Dim PT As POINTAPI
    GetCursorPos PT
    WC.ShowSysMenu PT.x, PT.y
End Sub


Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    'Do nothing
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


    If hwnd = TitleBarForm.hwnd Then
        If iMsg = WM_MOVE Then
            If ParentState <> vbMinimized Then
                Title.SetTopMost True
                MoveParent = True
                TitleBarForm.RefreshOnMove
                MoveParent = False
'               If ModalFlag = False Then Title.SetTopMost False
            End If
        End If
        
            
        
        If iMsg = WM_ACTIVATE Then
            If LoWord(wParam) = WA_ACTIVE Or LoWord(wParam) = WA_CLICKACTIVE Then
                If lParam <> Pfrm.hwnd Then
                    BringWindowToTop Me.hwnd
                    BringWindowToTop Pfrm.hwnd
                    Active = True
                End If
            Else
                If lParam = Pfrm.hwnd Then
                    Active = True
                Else
                    Active = False
                End If
            End If
            Call ChkGrad
        End If
    End If
    
    If hwnd = Pfrm.hwnd Then
        If iMsg = WM_STYLECHANGED Then
            Cap.Caption = Pfrm.Caption
        End If
        If iMsg = WM_SETICON Then FrmIcon.Picture = Pfrm.Icon
        
        
        If iMsg = WM_ACTIVATE Then
            If LoWord(wParam) = WA_ACTIVE Or LoWord(wParam) = WA_CLICKACTIVE Then
                Active = True
                If lParam <> TitleBarForm.hwnd Then
                    BringWindowToTop Me.hwnd
                    BringWindowToTop Pfrm.hwnd
                End If
            Else
                If lParam = TitleBarForm.hwnd Then
                    Active = True
                Else
                    Active = False
                End If
            End If
            Call ChkGrad
        End If
        
        If iMsg = WM_MOVE Then
            If MoveParent = False Then
                Call GetBorders
                Call CheckWinBtns
                Call GetBorders
                
                Call GetWindowRect(Pfrm.hwnd, pRect)
                If WC.ToolWindow = False Then
                    SetWindowPos Me.hwnd, 0&, pRect.Left, pRect.Top - (NCapSize + sBorder), pRect.Right - pRect.Left, (NCapSize + sBorder), SWP_NOACTIVATE
                Else
                    SetWindowPos Me.hwnd, 0&, pRect.Left, pRect.Top - (SCapSize + sBorder), pRect.Right - pRect.Left, (SCapSize + sBorder), SWP_NOACTIVATE
                End If
            End If
        End If
        
        If iMsg = WM_WINDOWPOSCHANGED Then
            If MoveParent = False Then
                If IsWindowVisible(Pfrm.hwnd) = 0 Then
                    If pVisible = True Then
                        Call ShowWindow(Me.hwnd, SW_HIDE)
                        pVisible = False
                    End If
                Else
                    If pVisible = False Then
                        Call ShowWindow(Me.hwnd, SW_SHOW)
                        pVisible = True
                    End If
                End If
            End If
        End If

        
        If iMsg = WM_APP + 1 Then
            SendMessage Pfrm.hwnd, WM_CLOSE, 0, 0
        End If
       
        If iMsg = WM_SHOWWINDOW Then
'           Call RefreshAll
        End If
    End If
End Function
     


Private Sub MaxRest_Click()
MaxRest.Clear
Call MaxRest.RestoreMaxBtn
    Select Case Pfrm.WindowState
        Case Is = vbMaximized
            Pfrm.WindowState = vbNormal
            MaxRest.Caption = "1"
            Title.SetAutoDrag True
        Case Is = vbNormal
            Pfrm.WindowState = vbMaximized
            MaxRest.Caption = "2"
            Title.SetAutoDrag False
    End Select
    Pfrm.SetFocus
End Sub

Private Sub Min_Click()
    Me.Hide
    Pfrm.SetFocus
    Pfrm.WindowState = vbMinimized
End Sub


Private Sub Pfrm_Resize()
    If Init = False Then Exit Sub
    If ByPassRefresh = True Then Exit Sub
    Call RefreshAll
    ParentState = Pfrm.WindowState
End Sub

Sub SetPicture(Pic As IPictureDisp)
    Set aPic = Pic
End Sub
Sub SetInactivePicture(Pic As IPictureDisp)
    Set inacPic = Pic
End Sub
Function Getpicture() As Picture
    Set Getpicture = Img.Picture
End Function


Private Sub Pfrm_Unload(Cancel As Integer)
    If Cancel <> True Then

        If ModalFlag = True Then EnableWindow Me.hwnd, False
        SetWindowLong Pfrm.hwnd, GWL_WNDPROC, GetOldProc()
        ByPassRefresh = True
        Init = False
        ByPassRefresh = True
        Unload Me
        
    End If
    
End Sub

Sub CheckWinBtns()
    CloseBtn.Enabled = WC.ControlBox
    MaxRest.Enabled = WC.MaxButton
    
    Min.Enabled = WC.MinButton
    FrmIcon.Visible = WC.ControlBox
End Sub


Sub SetFont(f As StdFont)
    Set Cap.Font = f
End Sub


Sub GetBorders()
    NCapSize = GetSystemMetrics(SM_CYCAPTION)
    SCapSize = GetSystemMetrics(SM_CYSMCAPTION)
   
    If WC.Sizable = True Then
        sBorder = GetSystemMetrics(SM_CYFRAME)
    Else
        sBorder = GetSystemMetrics(SM_CYDLGFRAME)
    End If
End Sub

Sub SETGStyle(Style As Integer)
GStyle = Style
End Sub

Function GETGStyle() As Integer
GETGStyle = GStyle
End Function

Sub RefreshOnMove()
If Pfrm.WindowState <> vbMaximized And Pfrm.WindowState <> vbMinimized Then
    If MoveParent = True Then
        Dim tRect As RECT
        GetWindowRect Me.hwnd, tRect
        GetWindowRect Pfrm.hwnd, pRect
        Call GetBorders
        If Title.ToolWindow = False Then
            SetWindowPos Pfrm.hwnd, 0&, tRect.Left, tRect.Top + NCapSize + sBorder, pRect.Right - pRect.Left, pRect.Bottom - pRect.Top, SWP_NOACTIVATE
        Else
            SetWindowPos Pfrm.hwnd, 0&, tRect.Left, tRect.Top + SCapSize + sBorder, pRect.Right - pRect.Left, pRect.Bottom - pRect.Top, SWP_NOACTIVATE
        End If
    End If
End If
End Sub


Function ChkMenu() As Boolean
    If Pfrm.BorderStyle <> 1 And Pfrm.BorderStyle <> 2 Then
        ChkMenu = False
        Exit Function
    End If
    If IsMenu(GetMenu(Pfrm.hwnd)) <> 0 Then ChkMenu = True
End Function

Sub SetInactiveColors(Color1 As OLE_COLOR, Color2 As OLE_COLOR)
IC1 = Color1
IC2 = Color2
Ing.GradColor1 = IC1
Ing.GradColor2 = IC2
End Sub

Sub InactiveGradient()
Me.Cls
    InacGrad = True
        Select Case GStyle
            Case Is = 1
                Ing.DoGradientHorizontal Me
            Case Is = 2
                Ing.DoGradientVertical Me
        End Select
End Sub

Sub ActiveGradient()
        Dim GC1 As OLE_COLOR, GC2 As OLE_COLOR
            InacGrad = False
            GC1 = GF.GradColor1
            GC2 = GF.GradColor2
            Me.Cls
            
                Select Case GStyle
                    Case Is = 1
                        GF.DoGradientHorizontal Me
                    Case Is = 2
                        GF.DoGradientVertical Me
                End Select

End Sub

Sub ChkGrad()
If Active = True Then
    Init = True
    Title.SetTopMost True
    Set Img.Picture = aPic
    If InacGrad = True Then
        ActiveGradient
        InacGrad = False
        SetWindowPos Pfrm.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_NOMOVE
        Cap.ForeColor = AcForeColor
'       If ModalFlag = True Then Title.SetTopMost True
'       If GetActiveWindow = Me.hwnd Then SetActiveWindow Pfrm.hwnd
    End If
Else
    If InacGrad = False Then
        For I = 1 To 10
            ByPassRefresh = True
            Title.SetTopMost False
            SetWindowPos Me.hwnd, GetForegroundWindow, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_NOMOVE
            Title.SetTopMost False
        Next I
        Set Img.Picture = inacPic
        InactiveGradient
        Cap.ForeColor = InTxt
        InacGrad = True
'       If ModalFlag = True Then Title.SetTopMost False
    End If
    ByPassRefresh = False
End If
End Sub

Sub SetInacTxtColor(Color As OLE_COLOR)
InTxt = Color
End Sub

Sub SetForeColor(Color As OLE_COLOR)
AcForeColor = Color
End Sub

