VERSION 5.00
Begin VB.Form SetLanguage 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form2"
   ScaleHeight     =   3855
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Language"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5775
      Begin VB.OptionButton Option2 
         Caption         =   "English - ÇäÌáíÒí"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Arabic - ÚÑÈí"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   960
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Choose main language , between Arabic or English ."
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   5295
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   0
      Picture         =   "SetLanguage.frx":0000
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
         Caption         =   "Select Language"
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
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   6000
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      X1              =   6000
      X2              =   6000
      Y1              =   3840
      Y2              =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Note : For good working , you should set screen resolution with 1024 * 768 ."
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   5535
   End
End
Attribute VB_Name = "SetLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RefreshLang()
Command1.Caption = LanguageTx(0)
Command2.Caption = LanguageTx(1)
Frame1.Caption = LanguageTx(2)
Label3.Caption = LanguageTx(3)
Label2.Caption = LanguageTx(4)
Label4.Caption = LanguageTx(5)
End Sub

Private Sub Command1_Click()
Open App.Path + "\Data\Languages\Selected.Txt" For Output As #1
If Option1.Value = True Then
Write #1, "Arb"
Else
Write #1, "Eng"
End If
Close #1
DoEvents

Main.Show
Unload Me
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Open App.Path + "\Data\Languages\Selected.Txt" For Input As #2
Input #2, x
If x = "Arb" Then
LoadLanguage True
Else
LoadLanguage False
Option2.Value = True
End If
Close #2
DoEvents

RefreshLang
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Option1_Click()
LoadLanguage True
RefreshLang
End Sub

Private Sub Option2_Click()
LoadLanguage False
RefreshLang
End Sub
