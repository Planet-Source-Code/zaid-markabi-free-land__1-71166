VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   ScaleHeight     =   4710
   ScaleMode       =   0  'User
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   2040
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   0
      Picture         =   "Splash.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Splash"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   50
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5640
         TabIndex        =   1
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Written with : Visual Basic 6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3840
      Width           =   6015
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Email : ZaidMarkabi@yahoo.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Work on : XP , Me , 2000"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Using : DirectX 8.0 with 3D MeshesSketch SDK Engine"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Free Land - 3D Land Generator"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Programmer : ZaidMarkabi"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading >>>"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   5535
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      X1              =   11.172
      X2              =   16.758
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   11.172
      X2              =   268.13
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   0
      Picture         =   "Splash.frx":6762
      Top             =   360
      Width           =   5985
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      X1              =   279.302
      X2              =   279.302
      Y1              =   4680
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   279.302
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   4680
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByValcrKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Const LWA_ALPHA = 2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000

Dim IntMe As Integer

Private Sub Form_Load()
   SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED

If App.PrevInstance = True Then
MsgBox "áÇ íãßä ÊÔÛíá ÃßËÑ ãä äÓÎÉ ãä ÇáÈÑäÇãÌ"
Unload Me
Exit Sub
End If
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If IntMe < 250 Then
IntMe = IntMe + 3
   SetLayeredWindowAttributes hwnd, 0, IntMe, LWA_ALPHA

Line5.x2 = Line5.x2 + Int(Rnd * 6.8)

Else
Timer1.Enabled = False
If Label3.Visible = True Then
SetLanguage.Show
IntMe = 0
Unload Me
End If
End If
End Sub
