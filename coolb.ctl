VERSION 5.00
Begin VB.UserControl CoolCommand 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   ClipControls    =   0   'False
   PropertyPages   =   "coolb.ctx":0000
   ScaleHeight     =   645
   ScaleWidth      =   1830
   ToolboxBitmap   =   "coolb.ctx":0035
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Command"
      Height          =   195
      Left            =   435
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Image Ico 
      Enabled         =   0   'False
      Height          =   510
      Left            =   120
      Stretch         =   -1  'True
      Top             =   75
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   975
      Left            =   30
      Top             =   60
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000F&
      Height          =   615
      Left            =   -75
      Top             =   -15
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000F&
      Height          =   600
      Left            =   60
      Top             =   165
      Width           =   1785
   End
End
Attribute VB_Name = "CoolCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'CoolCommand Control originally created by Marko Oette
'Extra GUI and GradientFX programming by
'Srideep Prasad,Pune city,Maharashtra,India

'Icon Support Added on 12Th June 2001
'Gradient Enable / Diable Bug Fix on 12Th June 2001

'NEW CODE DECLARATIONS
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
'END OF DECLARATION
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show

Private MMX As Variant
Private MMY As Variant
Private Highlighted As Boolean
Private Clicked As Boolean

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
'Default Property Values:
Const m_def_TextAlign = 0
Const m_def_FontHighlight = False
Const m_def_IconWidth = 32
Const m_def_IconHeight = 32
Const m_def_IconAlign = 0
Const m_def_IconSize = 2
Const m_def_Gradient = True
Const m_def_BackColor1 = &HE0E0E0
Const m_def_BackColor2 = &H404040
Const m_def_HighTextColor = &HFFFFFF
Const m_def_HighColor2 = &HFF8080
Const m_def_HighlightOnHover = True
Const m_def_HighlightStyle = 2
'Const m_def_CheckCaption = ""
Const m_def_HighColor1 = &HFFC0C0

Dim m_HighlightOnHover As Boolean
Dim m_HighlightStyle As Integer
Dim TColor As OLE_COLOR

'Dim m_CheckCaption As String
Dim m_HighColor1 As OLE_COLOR
Dim m_HighColor2 As OLE_COLOR
'Property Variables:
Dim m_TextAlign As Integer
Dim m_FontHighlight As Boolean
Dim m_IconWidth As Long
Dim m_IconHeight As Long
Dim m_IconAlign As Variant
Dim m_IconSize As Integer
Dim m_Gradient As Boolean
Dim m_BackColor1 As OLE_COLOR
Dim m_BackColor2 As OLE_COLOR
Dim m_HighTextColor As OLE_COLOR
Dim IconPic As IPictureDisp
Dim BoldFlag As Boolean


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Specifies the button forecolor"
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    TColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Specifies whether the control is enabled or not"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Specifies the font of the caption"
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    If Label1.FontBold = True Then
        BoldFlag = True
    Else
        BoldFlag = False
    End If
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    UserControl.Refresh
End Sub


Private Sub Label1_Change()
RefreshControl
End Sub



Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clicked = True
    BtnClick
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseMove(Button, Shift, x, y)
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clicked = False
    Highlight
End Sub

Private Function IsMouseOver() As Boolean
    Dim P As POINTAPI
    GetCursorPos P
    If WindowFromPoint(P.x, P.y) = UserControl.hwnd Then
        IsMouseOver = True
     Else
        IsMouseOver = False
    End If
End Function

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
RefreshControl
End Sub

Private Sub UserControl_ExitFocus()
UnHighlight
End Sub

Private Sub UserControl_Initialize()
UnHighlight
RefreshControl
        
        m_HighColor1 = m_def_HighColor1
        m_HighlightOnHover = m_def_HighlightOnHover
        m_HighlightStyle = m_def_HighlightStyle
        m_HighColor2 = m_def_HighColor2
        m_HighTextColor = m_def_HighTextColor
        m_BackColor1 = m_def_BackColor1
        m_BackColor2 = m_def_BackColor2
        
        UserControl.Cls
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent Click
End Sub

Private Sub UserControl_LostFocus()
UnHighlight
Shape3.Visible = False
End Sub
Private Sub Highlight()
If Highlighted = True Then Exit Sub
Shape1.BorderColor = &HE0E0E0
Shape2.BorderColor = &H808080
RefreshControl
Highlighted = True
End Sub

Private Sub UnHighlight()
If Highlighted = False Then Exit Sub
Shape1.BorderColor = vbButtonFace
Shape2.BorderColor = vbButtonFace
Highlighted = False
End Sub
Private Sub BtnClick()

Shape1.BorderColor = &H808080
Shape2.BorderColor = &HE0E0E0
Label1.Left = Label1.Left + 20
Label1.Top = Label1.Top + 20
UnHighlight
RaiseEvent Click
Call UserControl_MouseMove(0, 0, 0, 0)
Highlighted = False

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'The New Mousemove event
 
  If GetCapture() = UserControl.hwnd Then
    If ((x < 0) Or (x > UserControl.Width)) Or ((y < 0) Or (y > UserControl.Height)) Then
      ' if the mouse is outside the bounds of the control
      ' release the mouse and reset the backcolor
      Call ReleaseCapture
      UserControl.Cls
      If Clicked = False Then UnHighlight
        UserControl.Cls
        If BoldFlag = False Then
            If m_FontHighlight = True Then
                Label1.FontBold = False
            End If
        End If
                
                
        Select Case m_HighlightStyle
            Case Is = 2
                Call DoGradientVertical(m_BackColor1, m_BackColor2)
            Case Is = 1
                Call DoGradientHorizontal(m_BackColor1, m_BackColor2)
        End Select
                    
      Label1.ForeColor = TColor
    End If
    
  Else ' otherwise capture the mouse and change the backcolor of the control
    
    Call Highlight
    UserControl.Cls
    Call SetCapture(UserControl.hwnd)
    
    If BoldFlag = False Then
        If m_FontHighlight = True Then
            Label1.FontBold = True
        End If
    End If

    
    If m_HighlightOnHover = True Then
    Label1.ForeColor = m_HighTextColor
        Select Case m_HighlightStyle
            Case Is = 2
                Call DoGradientVertical(m_HighColor1, m_HighColor2)
            Case Is = 1
                Call DoGradientHorizontal(m_HighColor1, m_HighColor2)
        End Select
    End If
  End If
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Ico.Enabled = True
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    TColor = Label1.ForeColor
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    Label1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    Label1.Caption = PropBag.ReadProperty("checkcaption", "Test")
    RefreshControl
    m_HighColor1 = PropBag.ReadProperty("HighColor1", m_def_HighColor1)
    m_HighlightOnHover = PropBag.ReadProperty("HighlightOnHover", m_def_HighlightOnHover)
    m_HighlightStyle = PropBag.ReadProperty("HighlightStyle", m_def_HighlightStyle)
    m_HighColor2 = PropBag.ReadProperty("HighColor2", m_def_HighColor2)
    Label1.Caption = PropBag.ReadProperty("Caption", "Command")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_BevelWidth = PropBag.ReadProperty("BevelWidth", m_def_BevelWidth)
    m_HighTextColor = PropBag.ReadProperty("HighTextColor", m_def_HighTextColor)
    m_BackColor1 = PropBag.ReadProperty("BackColor1", m_def_BackColor1)
    m_BackColor2 = PropBag.ReadProperty("BackColor2", m_def_BackColor2)
    m_Gradient = PropBag.ReadProperty("Gradient", m_def_Gradient)
    Set Ico.Picture = PropBag.ReadProperty("Icon", Nothing)
    m_IconSize = PropBag.ReadProperty("IconSize", m_def_IconSize)
    m_IconAlign = PropBag.ReadProperty("IconAlign", m_def_IconAlign)
    Ico.Enabled = False
    m_IconWidth = PropBag.ReadProperty("IconWidth", m_def_IconWidth)
    m_IconHeight = PropBag.ReadProperty("IconHeight", m_def_IconHeight)
    m_FontHighlight = PropBag.ReadProperty("FontHighlight", m_def_FontHighlight)
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    Call UserControl_Resize
End Sub


Private Sub UserControl_Resize()
   Call AlignIcon
    RefreshControl
        UserControl.Cls
        If m_HighlightStyle <> 0 Then
            HighlightStyle = m_HighlightStyle
        Else
            HighlightStyle = m_def_HighlightStyle
        End If
        Select Case HighlightStyle
            Case Is = 2
                    Call DoGradientVertical(m_BackColor1, m_BackColor2)
            Case Is = 1
                    Call DoGradientHorizontal(m_BackColor1, m_BackColor2)
        End Select
            RaiseEvent Resize
    

End Sub
Sub RefreshControl()
Shape1.Top = 0
Shape2.Top = -10
Shape3.Top = 0
Shape1.Left = 0
Shape2.Left = -10
Shape3.Left = 0
Shape1.Width = Width + 10
Shape2.Width = Width + 10
Shape3.Width = Width
Shape1.Height = Height + 10
Shape2.Height = Height + 10
Shape3.Height = Height
Call AlignText
UnHighlight
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Ico.Enabled = True
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
    Call PropBag.WriteProperty("ToolTipText", Label1.ToolTipText, "")
    Call PropBag.WriteProperty("checkcaption", Label1.Caption, "Test")
    Call PropBag.WriteProperty("HighColor1", m_HighColor1, m_def_HighColor1)
    Call PropBag.WriteProperty("HighlightOnHover", m_HighlightOnHover, m_def_HighlightOnHover)
    Call PropBag.WriteProperty("HighlightStyle", m_HighlightStyle, m_def_HighlightStyle)
    Call PropBag.WriteProperty("HighColor2", m_HighColor2, m_def_HighColor2)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Command")
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)

    Call PropBag.WriteProperty("BevelWidth", m_BevelWidth, m_def_BevelWidth)
    Call PropBag.WriteProperty("HighTextColor", m_HighTextColor, m_def_HighTextColor)
    Call PropBag.WriteProperty("BackColor1", m_BackColor1, m_def_BackColor1)
    Call PropBag.WriteProperty("BackColor2", m_BackColor2, m_def_BackColor2)
    Call PropBag.WriteProperty("Gradient", m_Gradient, m_def_Gradient)
    Call PropBag.WriteProperty("Icon", Ico.Picture, Nothing)
    Call PropBag.WriteProperty("IconSize", m_IconSize, m_def_IconSize)
    Call PropBag.WriteProperty("IconAlign", m_IconAlign, m_def_IconAlign)
    
    Ico.Enabled = False
    Call PropBag.WriteProperty("IconWidth", m_IconWidth, m_def_IconWidth)
    Call PropBag.WriteProperty("IconHeight", m_IconHeight, m_def_IconHeight)
    Call PropBag.WriteProperty("FontHighlight", m_FontHighlight, m_def_FontHighlight)
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, m_def_TextAlign)
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, x, y)
Clicked = True
BtnClick
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    Highlight
    Clicked = False
End Sub

Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Specifies the ToolTip Text"
    ToolTipText = Label1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Label1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    RefreshControl
    m_HighColor1 = m_def_HighColor1
    m_HighlightOnHover = m_def_HighlightOnHover
    m_HighlightStyle = m_def_HighlightStyle
    m_HighColor2 = m_def_HighColor2
    m_HighTextColor = m_def_HighTextColor
    m_BackColor1 = m_def_BackColor1
    m_BackColor2 = m_def_BackColor2
    m_Gradient = m_def_Gradient
    m_IconSize = m_def_IconSize
    m_IconAlign = m_def_IconAlign
    m_IconWidth = m_def_IconWidth
    m_IconHeight = m_def_IconHeight
    m_FontHighlight = m_def_FontHighlight
    m_TextAlign = m_def_TextAlign
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00C0FFFF&
Public Property Get HighColor1() As OLE_COLOR
Attribute HighColor1.VB_Description = "The starting gradient highlight color"
    HighColor1 = m_HighColor1
End Property

Public Property Let HighColor1(ByVal New_HighColor1 As OLE_COLOR)
    m_HighColor1 = New_HighColor1
    PropertyChanged "HighColor1"
End Property
'


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



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get HighlightOnHover() As Boolean
Attribute HighlightOnHover.VB_Description = "Specifies wheter the button should highlight itself on Mouse Hover"
    HighlightOnHover = m_HighlightOnHover
End Property

Public Property Let HighlightOnHover(ByVal New_HighlightOnHover As Boolean)
    m_HighlightOnHover = New_HighlightOnHover
    PropertyChanged "HighlightOnHover"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,2
Public Property Get HighlightStyle() As Integer
Attribute HighlightStyle.VB_Description = "Specifies the Gradient Style (1=Horizontal [slower]) (2=Vertical [fast])"
    HighlightStyle = m_HighlightStyle
End Property

Public Property Let HighlightStyle(ByVal New_HighlightStyle As Integer)
    m_HighlightStyle = New_HighlightStyle
        UserControl.Cls
        Select Case HighlightStyle
            Case Is = 2
                    Call DoGradientVertical(m_BackColor1, m_BackColor2)
            Case Is = 1
                    Call DoGradientHorizontal(m_BackColor1, m_BackColor2)
            Case Else
                    Err.Raise vbObjectError + 1, , "Invalid HighlightStyle Value"
                    m_HighlightStyle = 1
        End Select

    PropertyChanged "HighlightStyle"
End Property
 
'Modified Code of the GradientFX Control
 
Friend Sub DoGradientHorizontal(m_GradColor1 As OLE_COLOR, m_GradColor2 As OLE_COLOR)
UserControl.Cls
If m_Gradient = False Then
    UserControl.Cls
    Exit Sub
End If

Dim FAC As Double, RFac As Double, GFac As Double, BFac As Double
Dim Red As Integer, Green As Integer, Blue As Integer
Dim RS As Integer, GS As Integer, BS As Integer
Dim RGB1 As Long, RGB2 As Long

Dim OD As Long
        RGB1 = TranslateColor(m_GradColor1)
        RS = UnRGB(RGB1, 0)
        GS = UnRGB(RGB1, 1)
        BS = UnRGB(RGB1, 2)
        
        RGB2 = TranslateColor(m_GradColor2)
        
        Red = UnRGB(RGB2, 0) - RS
        Green = UnRGB(RGB2, 1) - GS
        Blue = UnRGB(RGB2, 2) - BS
        
        RFac = Red / UserControl.ScaleWidth
        GFac = Green / UserControl.ScaleWidth
        BFac = Blue / UserControl.ScaleWidth
        
        FAC = 255 / UserControl.ScaleWidth
    For INTLOOP = 0 To UserControl.ScaleWidth
            UserControl.Line (INTLOOP, UserControl.ScaleLeft)-(INTLOOP, UserControl.Height), RGB(RS + (RFac * INTLOOP), GS + (GFac * INTLOOP), BS + (BFac * INTLOOP)), B
    Next INTLOOP
End Sub

Friend Sub DoGradientVertical(m_GradColor1 As OLE_COLOR, m_GradColor2 As OLE_COLOR)

UserControl.Cls

If m_Gradient = False Then
    UserControl.Cls
    Exit Sub
End If


Dim FAC As Double, RFac As Double, GFac As Double, BFac As Double
Dim Red As Integer, Green As Integer, Blue As Integer
Dim RS As Integer, GS As Integer, BS As Integer
Dim OD As Long
Dim RGB1 As Long, RGB2 As Long

        RGB1 = TranslateColor(m_GradColor1)
        RS = UnRGB(RGB1, 0)
        GS = UnRGB(RGB1, 1)
        BS = UnRGB(RGB1, 2)
        
        RGB2 = TranslateColor(m_GradColor2)
        
        Red = UnRGB(RGB2, 0) - RS
        Green = UnRGB(RGB2, 1) - GS
        Blue = UnRGB(RGB2, 2) - BS
        RFac = Red / UserControl.ScaleHeight
        GFac = Green / UserControl.ScaleHeight
        BFac = Blue / UserControl.ScaleHeight
    FAC = 255 / UserControl.ScaleHeight
    For INTLOOP = 0 To UserControl.ScaleHeight
             UserControl.Line (UserControl.ScaleLeft, INTLOOP)-(UserControl.Width, INTLOOP), RGB(RS + (RFac * INTLOOP), GS + (GFac * INTLOOP), BS + (BFac * INTLOOP)), B
    Next INTLOOP
End Sub
'End Of Modified Code of the GradientFX Control


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFF80&
Public Property Get HighColor2() As OLE_COLOR
Attribute HighColor2.VB_Description = "The ending Gradient Highlight color"
    HighColor2 = m_HighColor2
End Property

Public Property Let HighColor2(ByVal New_HighColor2 As OLE_COLOR)
    m_HighColor2 = New_HighColor2
    PropertyChanged "HighColor2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    Call AlignText
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    Set Label1.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    Label1.MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFFFF&
Public Property Get HighTextColor() As OLE_COLOR
Attribute HighTextColor.VB_Description = "Specifies the highlighted text color"
    HighTextColor = m_HighTextColor
End Property

Public Property Let HighTextColor(ByVal New_HighTextColor As OLE_COLOR)
    m_HighTextColor = New_HighTextColor
    PropertyChanged "HighTextColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H0000FFFF&
Public Property Get BackColor1() As OLE_COLOR
Attribute BackColor1.VB_Description = "Specifies the start gradient color"
    BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor1 As OLE_COLOR)
    m_BackColor1 = New_BackColor1
    UserControl.Cls
        Select Case m_HighlightStyle
            Case Is = 2
                Call DoGradientVertical(m_BackColor1, m_BackColor2)
            Case Is = 1
                Call DoGradientHorizontal(m_BackColor1, m_BackColor2)
        End Select
    
    PropertyChanged "BackColor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FF8080&
Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "Specifies the end gradient color"
    BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
    m_BackColor2 = New_BackColor2
        UserControl.Cls
        Select Case m_HighlightStyle
            Case Is = 2
                Call DoGradientVertical(m_BackColor1, m_BackColor2)
            Case Is = 1
                Call DoGradientHorizontal(m_BackColor1, m_BackColor2)
        End Select

    PropertyChanged "BackColor2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get Gradient() As Boolean
Attribute Gradient.VB_Description = "Enables/Disables the Gradient Effect"
    Gradient = m_Gradient
End Property

Public Property Let Gradient(ByVal New_Gradient As Boolean)
    m_Gradient = New_Gradient
    If m_Gradient = False Then
        UserControl.Cls
        PropertyChanged "Gradient"
        GoTo 50
    End If
    
    Select Case m_HighlightStyle
    Case Is = 1
        Call DoGradientHorizontal(m_BackColor1, m_BackColor2)
    Case Is = 2
        Call DoGradientVertical(m_BackColor1, m_BackColor2)
    End Select
    
    PropertyChanged "Gradient"
50 End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get IconSize() As Integer
Attribute IconSize.VB_Description = "Sprcifies the size of the icon (1=Small [16 x 16]) (2=Large [32 x 32]) (3=Custom) (4=Stretch to fill button)"
    IconSize = m_IconSize
End Property

Public Property Let IconSize(ByVal New_IconSize As Integer)
    m_IconSize = New_IconSize
    Call AlignIcon
    PropertyChanged "IconSize"
    PropertyChanged "IconHeight"
    PropertyChanged "IconWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get IconAlign() As Variant
Attribute IconAlign.VB_Description = "Specifies the Icon alignment (0=Center) (1=Left) (2=Right) (3=Top) (4=Bottom)"
    IconAlign = m_IconAlign
End Property

Public Property Let IconAlign(ByVal New_IconAlign As Variant)
    m_IconAlign = New_IconAlign
    Call AlignIcon
    PropertyChanged "IconAlign"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Ico,Ico,-1,Picture
Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Icon = Ico.Picture
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set Ico.Picture = New_Icon
    Set IconPic = New_Icon
    Call AlignIcon
    PropertyChanged "Icon"
End Property

Friend Sub AlignIcon()
Dim H As Long, W As Long
Select Case m_IconSize
    Case Is = 1
        H = 16 * Screen.TwipsPerPixelY
        W = 16 * Screen.TwipsPerPixelX
        m_IconHeight = 16
        m_IconWidth = 16

    Case Is = 2
        H = 32 * Screen.TwipsPerPixelY
        W = 32 * Screen.TwipsPerPixelX
        m_IconHeight = 32
        m_IconWidth = 32
    Case Is = 3
        H = m_IconHeight * Screen.TwipsPerPixelY
        W = m_IconWidth * Screen.TwipsPerPixelX
    Case Is = 4
        H = UserControl.ScaleHeight
        W = UserControl.ScaleWidth
        m_IconHeight = (H / Screen.TwipsPerPixelY)
        m_IconWidth = (W / Screen.TwipsPerPixelX)
        PropertyChanged "IconHeight"
        PropertyChanged "IconWidth"
End Select
Ico.Height = H
Ico.Width = W
Select Case m_IconAlign
    Case Is = 0
        Ico.Top = (UserControl.ScaleHeight - Ico.Height) / 2
        Ico.Left = (UserControl.ScaleWidth - Ico.Width) / 2
    Case Is = 1
        Ico.Top = (UserControl.ScaleHeight - Ico.Height) / 2
        Ico.Left = 0
    Case Is = 2
        Ico.Top = (UserControl.ScaleHeight - Ico.Height) / 2
        Ico.Left = (UserControl.ScaleWidth - Ico.Width)
    Case Is = 3
        Ico.Top = 0
        Ico.Left = (UserControl.ScaleWidth - Ico.Width) / 2
    Case Is = 4
        Ico.Top = UserControl.ScaleHeight - Ico.Height
        Ico.Left = (UserControl.ScaleWidth - Ico.Width) / 2
End Select
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,32
Public Property Get IconWidth() As Long
Attribute IconWidth.VB_Description = "Specifies the Icon Width"
    IconWidth = m_IconWidth
End Property

Public Property Let IconWidth(ByVal New_IconWidth As Long)
    m_IconWidth = New_IconWidth
    m_IconSize = 3
    Call AlignIcon
    PropertyChanged "IconWidth"
    PropertyChanged "IconSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,32
Public Property Get IconHeight() As Long
Attribute IconHeight.VB_Description = "Specifies the height of the icon"
    IconHeight = m_IconHeight
End Property

Public Property Let IconHeight(ByVal New_IconHeight As Long)
    m_IconHeight = New_IconHeight
    m_IconSize = 3
    Call AlignIcon
    PropertyChanged "IconHeight"
    PropertyChanged "IconSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get FontHighlight() As Boolean
Attribute FontHighlight.VB_Description = "Boldens the font on mouse hover"
    FontHighlight = m_FontHighlight
End Property

Public Property Let FontHighlight(ByVal New_FontHighlight As Boolean)
    m_FontHighlight = New_FontHighlight
    PropertyChanged "FontHighlight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get TextAlign() As Integer
Attribute TextAlign.VB_Description = "Sets the text alignment (0=Center) (1=Left) (2=Right) (3=Top) (4=Bottom)"
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal New_TextAlign As Integer)
    m_TextAlign = New_TextAlign
    Call AlignText
    PropertyChanged "TextAlign"
End Property

Friend Sub AlignText()
Select Case m_TextAlign
    Case Is = 0
        Label1.Left = (UserControl.ScaleWidth - Label1.Width) / 2
        Label1.Top = (UserControl.ScaleHeight - Label1.Height) / 2
    Case Is = 1
        Label1.Left = 0
        Label1.Top = (UserControl.ScaleHeight - Label1.Height) / 2
    Case Is = 2
        Label1.Left = (UserControl.ScaleWidth - Label1.Width)
        Label1.Top = (UserControl.ScaleHeight - Label1.Height) / 2
    Case Is = 3
        Label1.Top = 0
        Label1.Left = (UserControl.ScaleWidth - Label1.Width) / 2
    Case Is = 4
        Label1.Top = UserControl.ScaleHeight - Label1.Height
        Label1.Left = (UserControl.ScaleWidth - Label1.Width) / 2
End Select

End Sub

Sub RestoreMaxBtn()
      Call ReleaseCapture
      UserControl.Cls
      If Clicked = False Then UnHighlight
        UserControl.Cls
        Select Case m_HighlightStyle
            Case Is = 2
                Call DoGradientVertical(m_BackColor1, m_BackColor2)
            Case Is = 1
                Call DoGradientHorizontal(m_BackColor1, m_BackColor2)
        End Select

      Label1.ForeColor = TColor
End Sub
Sub Clear()
    UserControl.Cls
End Sub

