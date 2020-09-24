VERSION 5.00
Begin VB.Form SetLightOpt 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2745
      TabIndex        =   19
      Text            =   "0"
      Top             =   4320
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "SetLightOpt.frx":0000
      Left            =   2160
      List            =   "SetLightOpt.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3840
      Width           =   2415
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
      Left            =   2760
      TabIndex        =   15
      Text            =   "0.05"
      Top             =   3360
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
      Left            =   2760
      TabIndex        =   14
      Text            =   "0.1"
      Top             =   2880
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
      Left            =   2760
      TabIndex        =   13
      Text            =   "0.1"
      Top             =   2520
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
      Left            =   2760
      TabIndex        =   12
      Text            =   "0.1"
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Light Enable"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   255
      TabIndex        =   4
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1815
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
      Left            =   2760
      TabIndex        =   3
      Text            =   "2"
      Top             =   1800
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   0
      Picture         =   "SetLightOpt.frx":0053
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
         Caption         =   "Set Light Options"
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
   Begin VB.Label Label11 
      Caption         =   "Seconds"
      Height          =   255
      Left            =   3600
      TabIndex        =   21
      Top             =   4370
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Light Rotation :"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Light Direction :"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Shadow Power :"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Blue"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Green"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Red"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Light Color :"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   100
      Picture         =   "SetLightOpt.frx":67B5
      Top             =   480
      Width           =   5805
   End
   Begin VB.Label Label4 
      Caption         =   "Light Power :"
      Height          =   255
      Left            =   255
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
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
Attribute VB_Name = "SetLightOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub RefreshLang()
Label2.Caption = LanguageTx(54)
Check1.Caption = LanguageTx(55)
Label4.Caption = LanguageTx(56)
Label3.Caption = LanguageTx(57)
Label10.Caption = LanguageTx(58)
Label9.Caption = LanguageTx(59)
Label5.Caption = LanguageTx(60)
Label6.Caption = LanguageTx(61)
Label7.Caption = LanguageTx(62)
Combo1.List(0) = LanguageTx(63)
Combo1.List(1) = LanguageTx(64)
Combo1.List(2) = LanguageTx(65)
Combo1.List(3) = LanguageTx(66)
Combo1.List(4) = LanguageTx(67)
Label8.Caption = LanguageTx(87)
Label11.Caption = LanguageTx(88)

Command2.Caption = LanguageTx(18)
Command1.Caption = LanguageTx(19)
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Open App.Path + "\Data\Options\LightOptions.Cnf" For Output As #1
Write #1, Check1.Value
Write #1, Text2.Text
Write #1, Text1.Text
Write #1, Text3.Text
Write #1, Text4.Text
Write #1, Text5.Text
Write #1, Combo1.ListIndex
Write #1, Text6.Text

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

Open App.Path + "\Data\Options\LightOptions.Cnf" For Input As #1
Input #1, x
Check1.Value = x
Input #1, x
Text2.Text = x
Input #1, x
Text1.Text = x
Input #1, x
Text3.Text = x
Input #1, x
Text4.Text = x
Input #1, x
Text5.Text = x
Input #1, x
Combo1.ListIndex = Int(x)
Input #1, x
Text6.Text = x
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
