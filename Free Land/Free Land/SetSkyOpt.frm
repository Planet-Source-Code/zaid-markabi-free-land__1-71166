VERSION 5.00
Begin VB.Form SetSkyOpt 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3000
      TabIndex        =   19
      Text            =   "0.75"
      Top             =   4440
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
      Left            =   2160
      TabIndex        =   17
      Text            =   "0.5"
      Top             =   4440
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
      Left            =   2160
      TabIndex        =   15
      Text            =   "120"
      Top             =   4080
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      Caption         =   "Box Sky"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "Sphere Sky"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "SetSkyOpt.frx":0000
      Left            =   2160
      List            =   "SetSkyOpt.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
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
      Left            =   2160
      TabIndex        =   9
      Text            =   "32"
      Top             =   3240
      Width           =   735
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
      Left            =   4080
      TabIndex        =   7
      Text            =   "0.01"
      Top             =   2850
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      Caption         =   "Sky Rotation"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Sky Enable"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   5040
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   0
      Picture         =   "SetSkyOpt.frx":0038
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
         Caption         =   "Set Sky Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   50
         Width           =   5415
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   240
      X2              =   5760
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   240
      X2              =   5760
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label6 
      Caption         =   "Value Of Cloudes :"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Level Of Cloudes :"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Sky Mode :"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Set Poly Count :"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Set Speed ( by Radian ) :"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   100
      Picture         =   "SetSkyOpt.frx":679A
      Top             =   480
      Width           =   5820
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   5640
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   6000
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      X1              =   6000
      X2              =   6000
      Y1              =   5640
      Y2              =   240
   End
End
Attribute VB_Name = "SetSkyOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub RefreshLang()
Label2.Caption = LanguageTx(12)
Check1.Caption = LanguageTx(13)
Check2.Caption = LanguageTx(15)
Label4.Caption = LanguageTx(16)
Label3.Caption = LanguageTx(17)
Command2.Caption = LanguageTx(18)
Command1.Caption = LanguageTx(19)
Label7.Caption = LanguageTx(114)
Combo1.List(0) = LanguageTx(115)
Combo1.List(1) = LanguageTx(116)
Combo1.List(2) = LanguageTx(117)
Combo1.List(3) = LanguageTx(118)
Option1.Caption = LanguageTx(119)
Option2.Caption = LanguageTx(120)
Label5.Caption = LanguageTx(121)
Label6.Caption = LanguageTx(122)

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Open App.Path + "\Data\Options\SkyOptions.Cnf" For Output As #1
Write #1, Check1.Value
Write #1, Text2.Text
Write #1, Check2.Value
Write #1, Text1.Text
Write #1, Combo1.ListIndex
Write #1, Option1.Value
Write #1, Text3.Text
Write #1, Text4.Text
Write #1, Text5.Text

Close #1

Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Option1.Value = True

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

Open App.Path + "\Data\Options\SkyOptions.Cnf" For Input As #1
Input #1, x
Check1.Value = x
Input #1, x
Text2.Text = x
Input #1, x
Check2.Value = x
Input #1, x
Text1.Text = x
Input #1, x
Combo1.ListIndex = Int(x)
Input #1, x
If x = True Then
Option1.Value = True
Option2.Value = False
Else
Option2.Value = True
Option1.Value = False
End If
Input #1, x
Text3.Text = x
Input #1, x
Text4.Text = x
Input #1, x
Text5.Text = x
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
