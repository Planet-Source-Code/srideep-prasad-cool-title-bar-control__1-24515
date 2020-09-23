VERSION 5.00
Begin VB.UserControl CoolCaptionMDI 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   585
   InvisibleAtRuntime=   -1  'True
   Picture         =   "CoolCaptionMDI.ctx":0000
   PropertyPages   =   "CoolCaptionMDI.ctx":0442
   ScaleHeight     =   540
   ScaleWidth      =   585
   ToolboxBitmap   =   "CoolCaptionMDI.ctx":0468
   Begin VB.Image Image1 
      Height          =   510
      Left            =   15
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "CoolCaptionMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const GWL_WNDPROC = (-4)


Dim Par As MDIForm, CFrm As CapFrmMDI
Const m_def_BackColor1 = &HFF8080
Const m_def_BackColor2 = &H800000
Const m_def_CaptionGradStyle = 2
Const m_def_BtnForecolor = 0
Const m_def_HighlightStyle = 2
Const m_def_BtnBack1 = &HFFFFFF
Const m_def_BtnBack2 = &H404040
Const m_def_BtnHigh1 = &HFFFFFF
Const m_def_BtnHigh2 = &HFF8080
Const m_def_BtnHighText = &H800000
Const m_def_ForeColor = &HFFFFFF
Dim m_ForeColor As OLE_COLOR
Dim m_Picture As Picture
Dim m_BackColor1 As OLE_COLOR
Dim m_BackColor2 As OLE_COLOR
Dim m_CaptionGradStyle As Integer
Dim m_BtnForecolor As OLE_COLOR
Dim m_HighlightStyle As Integer
Dim m_BtnBack1 As OLE_COLOR
Dim m_BtnBack2 As OLE_COLOR
Dim m_BtnHigh1 As OLE_COLOR
Dim m_BtnHigh2 As OLE_COLOR
Dim m_BtnHighText As OLE_COLOR
Enum GStylesMDI
    grHorizontal = 1
    grVertical = 2
End Enum
'Default Property Values:
Const m_def_InacTxtColor = &HC0C0C0
Const m_def_InacColor1 = &H808080
Const m_def_InacColor2 = &H404040
'Property Variables:
Dim m_InactivePicture As Picture
Dim m_InacTxtColor As OLE_COLOR
Dim m_InacColor1 As OLE_COLOR
Dim m_InacColor2 As OLE_COLOR




Sub Init(Optional ParentForm As Object)
If ParentForm Is Nothing Then
    Set ParentForm = UserControl.Parent
End If
    
    Set Par = ParentForm

Set CFrm = New CapFrmMDI
If GetProp(Par.hwnd, "init") = 1 Then
    Err.Raise vbObjectError + 1, , "Form already coolcaptionized !"
    Exit Sub
End If

Load CFrm
CFrm.SetParent Par
CFrm.Show


        CFrm.GF.GradColor1 = m_BackColor1
        CFrm.GF.GradColor2 = m_GradColor2
    Select Case m_CaptionGradStyle
        Case Is = 1
            CFrm.Cls
            CFrm.GF.DoGradientHorizontal CFrm
        Case Is = 2
            CFrm.Cls
            CFrm.GF.DoGradientVertical CFrm
        Case Else
            CFrm.Cls
            MsgBox "Invalid gradient style"
    End Select
CFrm.Cap.ForeColor = m_ForeColor
Call CFrm.SetPicture(m_Picture)
SetProp Par.hwnd, "init", 1
Call Update
End Sub

Sub UnInit()
Unload CFrm
RemoveProp Par.hwnd, "init"
Set CFrm = Nothing
Set Par = Nothing
Set TitleBarForm = Nothing
InitFlag = False
End Sub
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=10,0,0,0
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = m_BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    m_BackColor = New_BackColor
'    Call Update
'    PropertyChanged "BackColor"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = m_Picture
    
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    Call Update
    PropertyChanged "Picture"
End Property




'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_BackColor = m_def_BackColor
    Set m_Picture = LoadPicture("")
    Set m_Font = Ambient.Font
    m_ForeColor = m_def_ForeColor
'    m_AutoUpdateInterval = m_def_AutoUpdateInterval
    m_BtnBack1 = m_def_BtnBack1
    m_BtnBack2 = m_def_BtnBack2
    m_BtnHigh1 = m_def_BtnHigh1
    m_BtnHigh2 = m_def_BtnHigh2
    m_BtnHighText = m_def_BtnHighText
    m_HighlightStyle = m_def_HighlightStyle
'    m_AutoUpdateInterval = m_def_AutoUpdateInterval
    m_BtnForecolor = m_def_BtnForecolor
    m_BackColor1 = m_def_BackColor1
    m_BackColor2 = m_def_BackColor2
    m_CaptionGradStyle = m_def_CaptionGradStyle
'    m_Caption = m_def_Caption
'    Set m_Icon = LoadPicture("")
'    m_AutoUpdateInterval = m_def_AutoUpdateInterval
    m_InacColor1 = m_def_InacColor1
    m_InacColor2 = m_def_InacColor2
    m_InacTxtColor = m_def_InacTxtColor
    Set m_InactivePicture = LoadPicture("")
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
'    m_AutoUpdateInterval = PropBag.ReadProperty("AutoUpdateInterval", m_def_AutoUpdateInterval)
    m_BtnBack1 = PropBag.ReadProperty("BtnBack1", m_def_BtnBack1)
    m_BtnBack2 = PropBag.ReadProperty("BtnBack2", m_def_BtnBack2)
    m_BtnHigh1 = PropBag.ReadProperty("BtnHigh1", m_def_BtnHigh1)
    m_BtnHigh2 = PropBag.ReadProperty("BtnHigh2", m_def_BtnHigh2)
    m_BtnHighText = PropBag.ReadProperty("BtnHighText", m_def_BtnHighText)
    m_HighlightStyle = PropBag.ReadProperty("HighlightStyle", m_def_HighlightStyle)
'    m_AutoUpdateInterval = PropBag.ReadProperty("AutoUpdateInterval", m_def_AutoUpdateInterval)
    m_BtnForecolor = PropBag.ReadProperty("BtnForecolor", m_def_BtnForecolor)
    m_BackColor1 = PropBag.ReadProperty("BackColor1", m_def_BackColor1)
    m_BackColor2 = PropBag.ReadProperty("BackColor2", m_def_BackColor2)
    m_CaptionGradStyle = PropBag.ReadProperty("CaptionGradStyle", m_def_CaptionGradStyle)
'    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
'    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
'    m_AutoUpdateInterval = PropBag.ReadProperty("AutoUpdateInterval", m_def_AutoUpdateInterval)
    m_InacColor1 = PropBag.ReadProperty("InacColor1", m_def_InacColor1)
    m_InacColor2 = PropBag.ReadProperty("InacColor2", m_def_InacColor2)
    m_InacTxtColor = PropBag.ReadProperty("InacTxtColor", m_def_InacTxtColor)
    Set m_InactivePicture = PropBag.ReadProperty("InactivePicture", Nothing)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = Image1.Width
    UserControl.Height = Image1.Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
'    Call PropBag.WriteProperty("AutoUpdateInterval", m_AutoUpdateInterval, m_def_AutoUpdateInterval)
    Call PropBag.WriteProperty("BtnBack1", m_BtnBack1, m_def_BtnBack1)
    Call PropBag.WriteProperty("BtnBack2", m_BtnBack2, m_def_BtnBack2)
    Call PropBag.WriteProperty("BtnHigh1", m_BtnHigh1, m_def_BtnHigh1)
    Call PropBag.WriteProperty("BtnHigh2", m_BtnHigh2, m_def_BtnHigh2)
    Call PropBag.WriteProperty("BtnHighText", m_BtnHighText, m_def_BtnHighText)
    Call PropBag.WriteProperty("HighlightStyle", m_HighlightStyle, m_def_HighlightStyle)
'    Call PropBag.WriteProperty("AutoUpdateInterval", m_AutoUpdateInterval, m_def_AutoUpdateInterval)
    Call PropBag.WriteProperty("BtnForecolor", m_BtnForecolor, m_def_BtnForecolor)
    Call PropBag.WriteProperty("BackColor1", m_BackColor1, m_def_BackColor1)
    Call PropBag.WriteProperty("BackColor2", m_BackColor2, m_def_BackColor2)
    Call PropBag.WriteProperty("CaptionGradStyle", m_CaptionGradStyle, m_def_CaptionGradStyle)
'    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
'    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
'    Call PropBag.WriteProperty("AutoUpdateInterval", m_AutoUpdateInterval, m_def_AutoUpdateInterval)
    Call PropBag.WriteProperty("InacColor1", m_InacColor1, m_def_InacColor1)
    Call PropBag.WriteProperty("InacColor2", m_InacColor2, m_def_InacColor2)
    Call PropBag.WriteProperty("InacTxtColor", m_InacTxtColor, m_def_InacTxtColor)
    Call PropBag.WriteProperty("InactivePicture", m_InactivePicture, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    Call Update
    PropertyChanged "ForeColor"
End Property

Sub Update()
If Par Is Nothing Then
    'Do nothing
Else
If GetProp(Par.hwnd, "init") = 1 Then
    CFrm.SetInacTxtColor m_InacTxtColor
    CFrm.SetForeColor (m_ForeColor)
    CFrm.Cap.Caption = Par.Caption
    
    CFrm.Min.BackColor1 = m_BtnBack1
    CFrm.MaxRest.BackColor1 = m_BtnBack1
    CFrm.CloseBtn.BackColor1 = m_BtnBack1
  
    CFrm.Min.BackColor2 = m_BtnBack2
    CFrm.MaxRest.BackColor2 = m_BtnBack2
    CFrm.CloseBtn.BackColor2 = m_BtnBack2
    
    CFrm.Min.HighColor1 = m_BtnHigh1
    CFrm.MaxRest.HighColor1 = m_BtnHigh1
    CFrm.CloseBtn.HighColor1 = m_BtnHigh1
    
    CFrm.Min.HighColor2 = m_BtnHigh2
    CFrm.MaxRest.HighColor2 = m_BtnHigh2
    CFrm.CloseBtn.HighColor2 = m_BtnHigh2
    
    CFrm.Min.HighTextColor = m_BtnHighText
    CFrm.MaxRest.HighTextColor = m_BtnHighText
    CFrm.CloseBtn.HighTextColor = m_BtnHighText
    
    CFrm.CloseBtn.HighlightStyle = m_HighlightStyle
    CFrm.Min.HighlightStyle = m_HighlightStyle
    CFrm.MaxRest.HighlightStyle = m_HighlightStyle
    
    CFrm.Min.ForeColor = m_BtnForecolor
    CFrm.MaxRest.ForeColor = m_BtnForecolor
    CFrm.CloseBtn.ForeColor = m_BtnForecolor
    If CFrm.GF.GradColor1 <> m_BackColor1 Or CFrm.GF.GradColor2 <> m_BackColor2 Or CFrm.GETGStyle <> m_CaptionGradStyle Then
        CFrm.GF.GradColor1 = m_BackColor1
        CFrm.GF.GradColor2 = m_BackColor2
    Select Case m_CaptionGradStyle
        Case Is = 1
            CFrm.Cls
            CFrm.GF.DoGradientHorizontal CFrm
        Case Is = 2
            CFrm.Cls
            CFrm.GF.DoGradientVertical CFrm
        Case Else
            CFrm.Cls
            MsgBox "Invalid gradient style"
    End Select
    End If
    CFrm.SetInactiveColors m_InacColor1, m_InacColor2
    CFrm.SETGStyle m_CaptionGradStyle
    CFrm.SetPicture m_Picture
    CFrm.SetInactivePicture m_InactivePicture
    CFrm.RefreshAll
End If
End If
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7,0,0,0
'Public Property Get AutoUpdateInterval() As Integer
'    AutoUpdateInterval = m_AutoUpdateInterval
'End Property
'
'Public Property Let AutoUpdateInterval(ByVal New_AutoUpdateInterval As Integer)
'    m_AutoUpdateInterval = New_AutoUpdateInterval
'    PropertyChanged "AutoUpdateInterval"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00E0E0E0&
Public Property Get BtnBack1() As OLE_COLOR
Attribute BtnBack1.VB_Description = "Sets/Returns the starting color of the Window Button Gradient Effect"
    BtnBack1 = m_BtnBack1
End Property

Public Property Let BtnBack1(ByVal New_BtnBack1 As OLE_COLOR)
    m_BtnBack1 = New_BtnBack1
    Call Update
    PropertyChanged "BtnBack1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00404040&
Public Property Get BtnBack2() As OLE_COLOR
Attribute BtnBack2.VB_Description = "Sets/Returns the ending color of the Window Button Gradient Effect"
    BtnBack2 = m_BtnBack2
End Property

Public Property Let BtnBack2(ByVal New_BtnBack2 As OLE_COLOR)
    m_BtnBack2 = New_BtnBack2
    Call Update
    PropertyChanged "BtnBack2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFF80&
Public Property Get BtnHigh1() As OLE_COLOR
Attribute BtnHigh1.VB_Description = "Sets/Returns the starting button highlight gradient color"
    BtnHigh1 = m_BtnHigh1
End Property

Public Property Let BtnHigh1(ByVal New_BtnHigh1 As OLE_COLOR)
    m_BtnHigh1 = New_BtnHigh1
    Call Update
    PropertyChanged "BtnHigh1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00C000C0&
Public Property Get BtnHigh2() As OLE_COLOR
Attribute BtnHigh2.VB_Description = "Sets/Returns the ending button highlight gradient color"
    BtnHigh2 = m_BtnHigh2
End Property

Public Property Let BtnHigh2(ByVal New_BtnHigh2 As OLE_COLOR)
    m_BtnHigh2 = New_BtnHigh2
    Call Update
    PropertyChanged "BtnHigh2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFFFF&
Public Property Get BtnHighText() As OLE_COLOR
Attribute BtnHighText.VB_Description = "Sets/Returns the color of the Window button symbols on highlight"
    BtnHighText = m_BtnHighText
End Property

Public Property Let BtnHighText(ByVal New_BtnHighText As OLE_COLOR)
    m_BtnHighText = New_BtnHighText
    Call Update
    PropertyChanged "BtnHighText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get ButtonGradStyle() As GStylesMDI
Attribute ButtonGradStyle.VB_Description = "Sets/Returns the Window button gradient style"
    ButtonGradStyle = m_HighlightStyle
End Property

Public Property Let ButtonGradStyle(ByVal New_ButtonGradStyle As GStylesMDI)
    m_HighlightStyle = New_ButtonGradStyle
    Call Update
    PropertyChanged "ButtonGradStyle"
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BtnForecolor() As OLE_COLOR
Attribute BtnForecolor.VB_Description = "Sets/Returns the color of the Window Button Symbols"
    BtnForecolor = m_BtnForecolor
End Property

Public Property Let BtnForecolor(ByVal New_BtnForecolor As OLE_COLOR)
    m_BtnForecolor = New_BtnForecolor
    Call Update
    PropertyChanged "BtnForecolor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor1() As OLE_COLOR
Attribute BackColor1.VB_Description = "Sets/Returns the starting color of the Active Window's Titlebar gradient effect"
    BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor1 As OLE_COLOR)
    m_BackColor1 = New_BackColor1
    Call Update
    PropertyChanged "BackColor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "Sets/Returns the ending color of the Active Window's Titlebar gradient effect"
    BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
    m_BackColor2 = New_BackColor2
    PropertyChanged "BackColor2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get CaptionGradStyle() As GStylesMDI
Attribute CaptionGradStyle.VB_Description = "Sets/Returns the gradient style of the Title Bar"
    CaptionGradStyle = m_CaptionGradStyle
End Property

Public Property Let CaptionGradStyle(ByVal New_CaptionGradStyle As GStylesMDI)
    m_CaptionGradStyle = New_CaptionGradStyle
    Call Update
    PropertyChanged "CaptionGradStyle"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00808080&
Public Property Get InacColor1() As OLE_COLOR
Attribute InacColor1.VB_Description = "Sets/Returns the starting gradient color or the inactive title bar"
    InacColor1 = m_InacColor1
End Property

Public Property Let InacColor1(ByVal New_InacColor1 As OLE_COLOR)
    m_InacColor1 = New_InacColor1
    Call Update
    PropertyChanged "InacColor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,&H00404040&
Public Property Get InacColor2() As OLE_COLOR
Attribute InacColor2.VB_Description = "Sets/Returns the ending gradient color or the inactive title bar"
    InacColor2 = m_InacColor2
End Property

Public Property Let InacColor2(ByVal New_InacColor2 As OLE_COLOR)
    m_InacColor2 = New_InacColor2
    Call Update
    PropertyChanged "InacColor2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00C0C0C0&
Public Property Get InacTxtColor() As OLE_COLOR
Attribute InacTxtColor.VB_Description = "Sets/Returns the text color of the inactive window's title bar"
    InacTxtColor = m_InacTxtColor
End Property

Public Property Let InacTxtColor(ByVal New_InacTxtColor As OLE_COLOR)
    m_InacTxtColor = New_InacTxtColor
    Call Update
    PropertyChanged "InacTxtColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get InactivePicture() As Picture
    Set InactivePicture = m_InactivePicture
End Property

Public Property Set InactivePicture(ByVal New_InactivePicture As Picture)
    Set m_InactivePicture = New_InactivePicture
    Call Update
    PropertyChanged "InactivePicture"
End Property

