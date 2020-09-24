VERSION 5.00
Begin VB.Form SetLandOpt 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   32
      Text            =   "1000"
      Top             =   3960
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      Caption         =   "Add Trees"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   31
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.ComboBox Combo4 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "SetLandOpt.frx":0000
      Left            =   2160
      List            =   "SetLandOpt.frx":0025
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   5520
      Width           =   2415
   End
   Begin VB.ComboBox Combo3 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "SetLandOpt.frx":0085
      Left            =   2160
      List            =   "SetLandOpt.frx":008F
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   5160
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "SetLandOpt.frx":009E
      Left            =   2160
      List            =   "SetLandOpt.frx":00AE
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   21
      Text            =   "3"
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   20
      Text            =   "3"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   19
      Text            =   "4"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   15
      Text            =   "4"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Text            =   "2"
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "SetLandOpt.frx":00D2
      Left            =   2160
      List            =   "SetLandOpt.frx":00E8
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Text            =   "10"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Text            =   "10"
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Terrain Enable"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   6120
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   0
      Picture         =   "SetLandOpt.frx":011D
      ScaleHeight     =   345
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5640
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Set Land Options"
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
         TabIndex        =   1
         Top             =   50
         Width           =   5415
      End
   End
   Begin VB.Line Line7 
      X1              =   5760
      X2              =   5760
      Y1              =   3120
      Y2              =   4440
   End
   Begin VB.Label Label17 
      Caption         =   "Distance Between Trees :"
      Height          =   255
      Left            =   3600
      TabIndex        =   33
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Line Line6 
      X1              =   3480
      X2              =   5760
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line5 
      X1              =   3480
      X2              =   5760
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line4 
      X1              =   3480
      X2              =   3480
      Y1              =   3120
      Y2              =   4440
   End
   Begin VB.Label Label16 
      Caption         =   "Global Texture :"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label15 
      Caption         =   "Detail Texture :"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "Detail Mode :"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "Details Num :"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "Y :"
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "X :"
      Height          =   255
      Left            =   2280
      TabIndex        =   22
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "X :"
      Height          =   255
      Left            =   2280
      TabIndex        =   18
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "Y :"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "Textures Num :"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Global Height :"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Land Quietly :"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Y :"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "X :"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Land Size :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   100
      Picture         =   "SetLandOpt.frx":687F
      Top             =   480
      Width           =   5805
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   6840
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   6000
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      X1              =   6000
      X2              =   6000
      Y1              =   6840
      Y2              =   240
   End
End
Attribute VB_Name = "SetLandOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub RefreshLang()
Label2.Caption = LanguageTx(25)
Command2.Caption = LanguageTx(18)
Command1.Caption = LanguageTx(19)
Label4.Caption = LanguageTx(26)
Label6.Caption = LanguageTx(27)
Combo1.List(0) = LanguageTx(28)
Combo1.List(1) = LanguageTx(29)
Combo1.List(2) = LanguageTx(30)
Combo1.List(3) = LanguageTx(31)
Combo1.List(4) = LanguageTx(32)
Combo1.List(5) = LanguageTx(33)
Label7.Caption = LanguageTx(34)
Label8.Caption = LanguageTx(35)
Label13.Caption = LanguageTx(36)
Label14.Caption = LanguageTx(37)
Label15.Caption = LanguageTx(38)
Label16.Caption = LanguageTx(39)
Combo2.List(0) = LanguageTx(40)
Combo2.List(1) = LanguageTx(41)
Combo2.List(2) = LanguageTx(42)
Combo2.List(3) = LanguageTx(43)
Combo3.List(0) = LanguageTx(44)
Combo3.List(1) = LanguageTx(45)
Combo4.List(0) = LanguageTx(46)
Combo4.List(1) = LanguageTx(47)
Combo4.List(2) = LanguageTx(48)
Combo4.List(3) = LanguageTx(49)
Combo4.List(4) = LanguageTx(50)
Combo4.List(5) = LanguageTx(51)
Combo4.List(6) = LanguageTx(52)
Check1.Caption = LanguageTx(53)
Combo4.List(7) = LanguageTx(90)
Combo4.List(8) = LanguageTx(91)
Combo4.List(9) = LanguageTx(92)
Combo4.List(10) = LanguageTx(93)
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Open App.Path + "\Data\Options\LandOptions.Cnf" For Output As #1
Write #1, Text3.Text
Write #1, Combo4.ListIndex
Write #1, Check1.Value
Write #1, Text2.Text
Write #1, Text1.Text
Write #1, Combo1.ListIndex
Close #1

Open App.Path + "\Data\Options\LandOptions2.Cnf" For Output As #1
Write #1, Combo3.ListIndex
Write #1, Text4.Text
Write #1, Text5.Text
Write #1, Text6.Text
Write #1, Text7.Text
Write #1, Combo2.ListIndex
Write #1, Check2.Value
Write #1, Text8.Text
Close #1

Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

Open App.Path + "\Data\Languages\Selected.Txt" For Input As #2
Input #2, x
If x = "Arb" Then
LoadLanguage True
Else
LoadLanguage False
End If
Close #2
DoEvents
RefreshLang

Open App.Path + "\Data\Options\LandOptions.Cnf" For Input As #1
Input #1, x
Text3.Text = x
Input #1, x
Combo4.ListIndex = Int(x)
Input #1, x
Check1.Value = x
Input #1, x
Text2.Text = x
Input #1, x
Text1.Text = x
Input #1, x
Combo1.ListIndex = Int(x)
Close #1

Open App.Path + "\Data\Options\LandOptions2.Cnf" For Input As #1
Input #1, x
Combo3.ListIndex = Int(x)
Input #1, x
Text4.Text = x
Input #1, x
Text5.Text = x
Input #1, x
Text6.Text = x
Input #1, x
Text7.Text = x
Input #1, x
Combo2.ListIndex = Int(x)
Input #1, x
Check2.Value = x
Input #1, x
Text8.Text = x
Close #1
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not KeyAscii = 46 Then
KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not KeyAscii = 46 Then
KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not KeyAscii = 46 Then
KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not KeyAscii = 46 Then
KeyAscii = 0
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not KeyAscii = 46 Then
KeyAscii = 0
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not KeyAscii = 46 Then
KeyAscii = 0
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not KeyAscii = 46 Then
KeyAscii = 0
End If
End Sub
