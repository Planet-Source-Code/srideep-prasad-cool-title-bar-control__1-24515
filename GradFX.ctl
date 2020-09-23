VERSION 5.00
Begin VB.UserControl GradientFX 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "GradFX.ctx":0000
   ScaleHeight     =   390
   ScaleWidth      =   465
   ToolboxBitmap   =   "GradFX.ctx":0014
   Begin VB.Image Image1 
      Height          =   480
      Left            =   15
      Picture         =   "GradFX.ctx":0326
      Top             =   -45
      Width           =   480
   End
End
Attribute VB_Name = "GradientFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Default Property Values:
Const m_def_DrawWidth = 1
Const m_def_AutoSetRedrawFlag = True
Const m_def_GradColor1 = &HFFFF&
Const m_def_GradColor2 = &HC00000
'Property Variables:
Dim m_DrawWidth As Integer
Dim m_AutoSetRedrawFlag As Boolean
Dim m_GradColor1 As OLE_COLOR
Dim m_GradColor2 As OLE_COLOR




Private Sub UserControl_Resize()
    UserControl.Width = Image1.Width
    UserControl.Height = Image1.Height
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get GradColor1() As OLE_COLOR
    GradColor1 = m_GradColor1
End Property

Public Property Let GradColor1(ByVal New_GradColor1 As OLE_COLOR)
    m_GradColor1 = New_GradColor1
    PropertyChanged "GradColor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get GradColor2() As OLE_COLOR
    GradColor2 = m_GradColor2
End Property

Public Property Let GradColor2(ByVal New_GradColor2 As OLE_COLOR)
    m_GradColor2 = New_GradColor2
    PropertyChanged "GradColor2"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_GradColor1 = m_def_GradColor1
    m_GradColor2 = m_def_GradColor2
'    m_Orientation = m_def_Orientation
    m_AutoSetRedrawFlag = m_def_AutoSetRedrawFlag
    m_DrawWidth = m_def_DrawWidth
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_GradColor1 = PropBag.ReadProperty("GradColor1", m_def_GradColor1)
    m_GradColor2 = PropBag.ReadProperty("GradColor2", m_def_GradColor2)
    m_AutoSetRedrawFlag = PropBag.ReadProperty("AutoSetRedrawFlag", m_def_AutoSetRedrawFlag)
    m_DrawWidth = PropBag.ReadProperty("DrawWidth", m_def_DrawWidth)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("GradColor1", m_GradColor1, m_def_GradColor1)
    Call PropBag.WriteProperty("GradColor2", m_GradColor2, m_def_GradColor2)
    Call PropBag.WriteProperty("AutoSetRedrawFlag", m_AutoSetRedrawFlag, m_def_AutoSetRedrawFlag)
    Call PropBag.WriteProperty("DrawWidth", m_DrawWidth, m_def_DrawWidth)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Sub DoGradientHorizontal(GObject As Object)
Dim FAC As Double, RFac As Double, GFac As Double, BFac As Double
Dim Red As Integer, Green As Integer, Blue As Integer
Dim RS As Integer, GS As Integer, BS As Integer
Dim RGB1 As Long, RGB2 As Long

Dim IPic As Object
Set IPic = GObject
Dim OD As Long
OD = IPic.DrawWidth
If m_AutoSetRedrawFlag = True Then IPic.AutoRedraw = True
    IPic.DrawWidth = m_DrawWidth
        RGB1 = TranslateColor(m_GradColor1)
        RS = UnRGB(RGB1, 0)
        GS = UnRGB(RGB1, 1)
        BS = UnRGB(RGB1, 2)
        
        RGB2 = TranslateColor(m_GradColor2)
        
        Red = UnRGB(RGB2, 0) - RS
        Green = UnRGB(RGB2, 1) - GS
        Blue = UnRGB(RGB2, 2) - BS
        
        RFac = Red / IPic.ScaleWidth
        GFac = Green / IPic.ScaleWidth
        BFac = Blue / IPic.ScaleWidth
        
        FAC = 255 / IPic.ScaleWidth
    For INTLOOP = 0 To IPic.ScaleWidth
            IPic.Line (INTLOOP, IPic.ScaleLeft)-(INTLOOP, IPic.Height), RGB(RS + (RFac * INTLOOP), GS + (GFac * INTLOOP), BS + (BFac * INTLOOP)), B
    Next INTLOOP
IPic.DrawWidth = OD
End Sub

Sub DoGradientVertical(GObject As Object)
Dim FAC As Double, RFac As Double, GFac As Double, BFac As Double
Dim Red As Integer, Green As Integer, Blue As Integer
Dim RS As Integer, GS As Integer, BS As Integer
Dim OD As Long
Dim RGB1 As Long, RGB2 As Long

Dim IPic As Object

Set IPic = GObject
OD = IPic.DrawWidth
If m_AutoSetRedrawFlag = True Then IPic.AutoRedraw = True
    IPic.DrawWidth = m_DrawWidth
        RGB1 = TranslateColor(m_GradColor1)
        RS = UnRGB(RGB1, 0)
        GS = UnRGB(RGB1, 1)
        BS = UnRGB(RGB1, 2)
        
        RGB2 = TranslateColor(m_GradColor2)
        
        Red = UnRGB(RGB2, 0) - RS
        Green = UnRGB(RGB2, 1) - GS
        Blue = UnRGB(RGB2, 2) - BS
        
        RFac = Red / IPic.ScaleHeight
        GFac = Green / IPic.ScaleHeight
        BFac = Blue / IPic.ScaleHeight
    FAC = 255 / IPic.ScaleHeight
    For INTLOOP = 0 To IPic.ScaleHeight
             IPic.Line (IPic.ScaleLeft, INTLOOP)-(IPic.Width, INTLOOP), RGB(RS + (RFac * INTLOOP), GS + (GFac * INTLOOP), BS + (BFac * INTLOOP)), B
    Next INTLOOP
IPic.DrawWidth = OD
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoSetRedrawFlag() As Boolean
    AutoSetRedrawFlag = m_AutoSetRedrawFlag
End Property

Public Property Let AutoSetRedrawFlag(ByVal New_AutoSetRedrawFlag As Boolean)
    m_AutoSetRedrawFlag = New_AutoSetRedrawFlag
    PropertyChanged "AutoSetRedrawFlag"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
    DrawWidth = m_DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    m_DrawWidth = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property

Sub DoPatternGradientV(GObject As Object, Optional Start As Long = 0, Optional GStep As Long = 1, Optional Width As Long = 1)
Dim FAC As Double, RFac As Double, GFac As Double, BFac As Double
Dim Red As Integer, Green As Integer, Blue As Integer
Dim RS As Integer, GS As Integer, BS As Integer
Dim OD As Long
Dim RGB1 As Long, RGB2 As Long

Dim IPic As Object

Set IPic = GObject
If m_AutoSetRedrawFlag = True Then IPic.AutoRedraw = True
    IPic.DrawWidth = m_DrawWidth
        RGB1 = TranslateColor(m_GradColor1)
        RS = UnRGB(RGB1, 0)
        GS = UnRGB(RGB1, 1)
        BS = UnRGB(RGB1, 2)
        
        RGB2 = TranslateColor(m_GradColor2)
        
        Red = UnRGB(RGB2, 0) - RS
        Green = UnRGB(RGB2, 1) - GS
        Blue = UnRGB(RGB2, 2) - BS
        
        RFac = (Red / (IPic.ScaleHeight - Start))
        GFac = (Green / (IPic.ScaleHeight - Start))
        BFac = (Blue / (IPic.ScaleHeight - Start))
        OD = IPic.DrawWidth
        If Width = 0 Then
            IPic.DrawWidth = Width
        Else
            IPic.DrawWidth = m_DrawWidth
        End If
            
    For INTLOOP = Start To GStep * Fix((IPic.ScaleHeight - Start) / GStep) Step GStep
             IPic.Line (IPic.ScaleLeft, INTLOOP)-(IPic.Width, INTLOOP), RGB(RS + (RFac * INTLOOP), GS + (GFac * INTLOOP), BS + (BFac * INTLOOP)), B
    Next INTLOOP
    IPic.DrawWidth = OD
End Sub


Sub DoPatternGradientH(GObject As Object, Optional Start As Long = 0, Optional GStep As Long = 1, Optional Width As Long = 1)
Dim FAC As Double, RFac As Double, GFac As Double, BFac As Double
Dim Red As Integer, Green As Integer, Blue As Integer
Dim RS As Integer, GS As Integer, BS As Integer
Dim OD As Long
Dim RGB1 As Long, RGB2 As Long

Dim IPic As Object

Set IPic = GObject
If m_AutoSetRedrawFlag = True Then IPic.AutoRedraw = True
    IPic.DrawWidth = m_DrawWidth
        RGB1 = TranslateColor(m_GradColor1)
        RS = UnRGB(RGB1, 0)
        GS = UnRGB(RGB1, 1)
        BS = UnRGB(RGB1, 2)
        
        RGB2 = TranslateColor(m_GradColor2)
        
        Red = UnRGB(RGB2, 0) - RS
        Green = UnRGB(RGB2, 1) - GS
        Blue = UnRGB(RGB2, 2) - BS
        
        RFac = Red / (IPic.ScaleWidth - Start)
        GFac = Green / (IPic.ScaleWidth - Start)
        BFac = Blue / (IPic.ScaleWidth - Start)
        OD = IPic.DrawWidth
        If Width = 0 Then
            IPic.DrawWidth = Width
        Else
            IPic.DrawWidth = m_DrawWidth
        End If
            
    FAC = 255 / ((IPic.ScaleWidth - Start) / GStep)
    For INTLOOP = Start To GStep * Fix((IPic.ScaleWidth - Start) / GStep) Step GStep
            IPic.Line (INTLOOP, IPic.ScaleLeft)-(INTLOOP, IPic.Height), RGB(RS + (RFac * INTLOOP), GS + (GFac * INTLOOP), BS + (BFac * INTLOOP)), B
    Next INTLOOP
    IPic.DrawWidth = OD
End Sub


Sub CheckedGradient(GObject As Object, Optional Start As Long = 0, Optional sqUnit As Long = 1, Optional Width As Long = 1)
Call DoPatternGradientV(GObject, Start, sqUnit, Width)
Call DoPatternGradientH(GObject, Start, sqUnit, Width)
End Sub
