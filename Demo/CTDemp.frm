VERSION 5.00
Object = "*\A..\CoolCaption.vbp"
Begin VB.Form CTDemp 
   BackColor       =   &H00000000&
   Caption         =   "Welcome !"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin CoolCaptionControl.CoolCaption C 
      Left            =   0
      Top             =   3960
      _ExtentX        =   979
      _ExtentY        =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Contact Me"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"CTDemp.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   7095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "So What are you waiting for ? Download it now !!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   6495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"CTDemp.frx":0093
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   3240
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A New Title Bar Control For VB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check out the new CoolCaption Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "CTDemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
MsgBox "Hope you liked it ! You could post any suggestions/opinions at planet-source-code.com (they will be e-mailed to me automatically) or you could e-mail me at srideepprasad@yahoo.com"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
C.Init Me
End Sub
