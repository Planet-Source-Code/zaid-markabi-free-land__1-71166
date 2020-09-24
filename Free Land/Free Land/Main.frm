VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Free Land"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15270
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
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
      Left            =   9000
      TabIndex        =   48
      Text            =   "60"
      Top             =   7440
      Width           =   975
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
      Left            =   9000
      TabIndex        =   47
      Text            =   "1000000"
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Apply"
      Height          =   375
      Left            =   10320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7680
      Width           =   1695
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
      Left            =   9000
      TabIndex        =   44
      Text            =   "9"
      Top             =   7080
      Width           =   975
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
      Left            =   9000
      TabIndex        =   42
      Text            =   "4"
      Top             =   6720
      Width           =   975
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
      Left            =   9000
      TabIndex        =   40
      Text            =   "0.04"
      Top             =   6360
      Width           =   975
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   6240
      Picture         =   "Main.frx":1CCA
      ScaleHeight     =   345
      ScaleWidth      =   6015
      TabIndex        =   38
      Top             =   5880
      Width           =   6015
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Camera Options"
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
         TabIndex        =   39
         Top             =   50
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   6240
      Picture         =   "Main.frx":1D10
      ScaleHeight     =   975
      ScaleWidth      =   6015
      TabIndex        =   34
      Top             =   4740
      Width           =   6015
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Global World"
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
         TabIndex        =   37
         Top             =   45
         Width           =   5415
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Load from File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2280
         TabIndex        =   36
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Save to File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   6240
      Picture         =   "Main.frx":152B2
      ScaleHeight     =   975
      ScaleWidth      =   6015
      TabIndex        =   29
      Top             =   3770
      Width           =   6015
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Save to File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Load from File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Map Edit"
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
         TabIndex        =   31
         Top             =   45
         Width           =   5415
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Height Map , Set Auto Map ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   6240
      Picture         =   "Main.frx":152F8
      ScaleHeight     =   975
      ScaleWidth      =   6015
      TabIndex        =   24
      Top             =   2790
      Width           =   6015
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Load from File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Save to File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Poly Count , Set Speed ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Sky Edit"
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
         TabIndex        =   25
         Top             =   45
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   6240
      Picture         =   "Main.frx":1533E
      ScaleHeight     =   975
      ScaleWidth      =   6015
      TabIndex        =   19
      Top             =   1820
      Width           =   6015
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Light Edit"
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
         TabIndex        =   23
         Top             =   45
         Width           =   5415
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Light Power , Shadow Power , Light Direction ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Save to File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Load from File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   6240
      Picture         =   "Main.frx":15384
      ScaleHeight     =   975
      ScaleWidth      =   6015
      TabIndex        =   14
      Top             =   840
      Width           =   6015
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Load from File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Save to File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Global Texture , Detail , Quietly , Size ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Land Edit"
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
         TabIndex        =   15
         Top             =   45
         Width           =   5415
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Set as Background"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   5775
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   6240
      Picture         =   "Main.frx":153CA
      ScaleHeight     =   345
      ScaleWidth      =   6015
      TabIndex        =   11
      Top             =   480
      Width           =   6015
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Options"
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
         TabIndex        =   12
         Top             =   50
         Width           =   5415
      End
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Main.frx":15410
      Left            =   1440
      List            =   "Main.frx":15420
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   7260
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Render 3D View"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   5775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear Image"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7680
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog SAV 
      Left            =   4680
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save as BMP File"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   5775
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   120
      Picture         =   "Main.frx":1544C
      ScaleHeight     =   345
      ScaleWidth      =   6015
      TabIndex        =   4
      Top             =   480
      Width           =   6015
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Image Rendered"
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
         TabIndex        =   5
         Top             =   50
         Width           =   5415
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   -60
      Picture         =   "Main.frx":1BBAE
      ScaleHeight     =   345
      ScaleWidth      =   15375
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin VB.Image Image2 
         Height          =   480
         Left            =   0
         Picture         =   "Main.frx":2C3F0
         Top             =   -80
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   14640
         TabIndex        =   3
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   15000
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Free Land"
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
         Left            =   600
         TabIndex        =   1
         Top             =   45
         Width           =   6975
      End
   End
   Begin VB.Image Image4 
      Height          =   7680
      Left            =   12360
      Picture         =   "Main.frx":2D0BA
      Top             =   480
      Width           =   2760
   End
   Begin VB.Image Image3 
      Height          =   2235
      Left            =   0
      Picture         =   "Main.frx":34762
      Top             =   8280
      Width           =   15360
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Camera Corner View :"
      Height          =   255
      Left            =   6480
      TabIndex        =   50
      Top             =   7440
      Width           =   2415
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Camera Max View :"
      Height          =   255
      Left            =   6480
      TabIndex        =   49
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Crackling While Walking :"
      Height          =   255
      Left            =   6480
      TabIndex        =   45
      Top             =   7080
      Width           =   2415
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Height of Eyes From The Land :"
      Height          =   255
      Left            =   6480
      TabIndex        =   43
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed Camera Walking :"
      Height          =   255
      Left            =   6480
      TabIndex        =   41
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Line Line10 
      BorderStyle     =   3  'Dot
      X1              =   12240
      X2              =   12240
      Y1              =   6000
      Y2              =   8160
   End
   Begin VB.Line Line9 
      BorderStyle     =   3  'Dot
      X1              =   6240
      X2              =   6240
      Y1              =   6000
      Y2              =   8160
   End
   Begin VB.Line Line8 
      BorderStyle     =   3  'Dot
      X1              =   12240
      X2              =   6240
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line6 
      BorderStyle     =   3  'Dot
      X1              =   6240
      X2              =   6240
      Y1              =   840
      Y2              =   5760
   End
   Begin VB.Line Line5 
      BorderStyle     =   3  'Dot
      X1              =   12240
      X2              =   12240
      Y1              =   840
      Y2              =   5760
   End
   Begin VB.Line Line4 
      BorderStyle     =   3  'Dot
      X1              =   12240
      X2              =   6240
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Size :"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   7290
      Width           =   1095
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3855
      Left            =   240
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      X1              =   6120
      X2              =   120
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   6120
      X2              =   6120
      Y1              =   840
      Y2              =   8160
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   120
      Y1              =   840
      Y2              =   8160
   End
   Begin VB.Menu WorldMnu 
      Caption         =   "World"
      Begin VB.Menu NewWorldMnu 
         Caption         =   "Edit Height Map"
      End
      Begin VB.Menu SPMnu1 
         Caption         =   "-"
      End
      Begin VB.Menu SaveBmpMnu 
         Caption         =   "Export Image as BMP File"
      End
      Begin VB.Menu SetBackMnu 
         Caption         =   "Set as Background"
      End
      Begin VB.Menu SPMnu2 
         Caption         =   "-"
      End
      Begin VB.Menu ExitSub 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu NTH1 
      Caption         =   "          "
   End
   Begin VB.Menu RenderMnu 
      Caption         =   "Render"
      Begin VB.Menu RenderViewMnu 
         Caption         =   "Render 3D View"
      End
      Begin VB.Menu ClsImgMnu 
         Caption         =   "Clear Rendered Image"
      End
   End
   Begin VB.Menu NTH2 
      Caption         =   "          "
   End
   Begin VB.Menu EditMnu 
      Caption         =   "Edit"
      Begin VB.Menu LandEditorMnu 
         Caption         =   "Land Editor"
      End
      Begin VB.Menu SkyEditorMnu 
         Caption         =   "Sky Editor"
      End
      Begin VB.Menu LightEditorMnu 
         Caption         =   "Light Editor"
      End
      Begin VB.Menu MapEditorMnu 
         Caption         =   "Map Editor"
      End
   End
   Begin VB.Menu NTH3 
      Caption         =   "          "
   End
   Begin VB.Menu AboutMnu 
      Caption         =   "About"
      Begin VB.Menu LndMnu 
         Caption         =   "About Free Land"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmed By [ Zaid Markabi ]
' ___________________________________________________________________________________________________
'|                                                                                                   |\_______________________
'|  ###############        ###         #####   ######                ######    #####                 |                        |\0 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1
'| ##############         #####         ###     ##   ##               ######  #####                  |      Zaid Markabi      |=\ 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0
'|         ####          ### ###        ###     ##    ##              ##  ## ##  ##                  |                        |==\0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1 0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1
'|       ###            ###   ###       ###     ##     ##    #####    ##   ###   ##                  | zaidmarkabi@yahoo.com  |===\ 1 __________________________________  0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 1 0 0 0 1
'|     ###             ###########      ###     ##     ##   ####      ##    #    ##                  |                        |====|>| Development For Our Digital Life | 1 1 0 0 1 1 1 0 1 0 0 1 0 0 0 1 1 0 1 0
'|   ###              #############     ###     ##    ##              ##         ##      A R K A B I | VisualBasic Programmer |===/ 1|__________________________________| 0 1 1 0 1 0 0 0 1 1 1 0 1 0 1 1 0 1 0 0
'| ##############    ###         ###    ###     ##   ##               ##         ##     ############ |                        |==/0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1 0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1
'| ###############   ###         ###   #####   ######                ####       ####   ### 2008 ###  |Syria(Arab Area)-Tartuse|=/ 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0
'|                                                                                    ############   | _______________________|/0 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1
'|___________________________________________________________________________________________________|/

' Email me at
' zaidmarkabi@yahoo.com
' I hope to hear from you soon,

' About Me
' --------
' Name:  Zaid Markabi
' Language:  Arabic
' Nationality:  Arabic - Syrian
' Live in : Syria - Tartuse
' My WebPage : http://yazanmarkabi.jeeran.com/allmembers/zaid%20markabi.html
' My Twin's Website : YazanMarkabi.Jeeran.com
' Email:  ZaidMarkabi@yahoo.com

Const SPI_SETDESKWALLPAPER = 20
Const SPIF_UPDATEINIFILE = &H1
Const SPIF_SENDWININICHANGE = &H2
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Sub RefreshLang()
Label2.Caption = LanguageTx(6)
Me.Caption = LanguageTx(6)
WorldMnu.Caption = LanguageTx(7)
NewWorldMnu.Caption = LanguageTx(8)
'OpenWorldMnu.Caption = LanguageTx(9)
'SaveWorldMnu.Caption = LanguageTx(10)
Label5.Caption = LanguageTx(68)
Command1.Caption = LanguageTx(69)
Command3.Caption = LanguageTx(70)
Command2.Caption = LanguageTx(71)
Label4.Caption = LanguageTx(72)
Combo1.List(3) = LanguageTx(73)
Label6.Caption = LanguageTx(74)
Label7.Caption = LanguageTx(75)
Label14.Caption = LanguageTx(76)
Label15.Caption = LanguageTx(77)
Label22.Caption = LanguageTx(78)
Label8.Caption = LanguageTx(79)
Label13.Caption = LanguageTx(80)
Label16.Caption = LanguageTx(81)
Label21.Caption = LanguageTx(82)
Label20.Caption = LanguageTx(83)
Label17.Caption = LanguageTx(83)
Label12.Caption = LanguageTx(83)
Label9.Caption = LanguageTx(83)
Label19.Caption = LanguageTx(84)
Label18.Caption = LanguageTx(84)
Label11.Caption = LanguageTx(84)
Label10.Caption = LanguageTx(84)
Command4.Caption = LanguageTx(85)
Label25.Caption = LanguageTx(89)
Label23.Caption = LanguageTx(83)
Label24.Caption = LanguageTx(84)
Label27.Caption = LanguageTx(94)
Label28.Caption = LanguageTx(95)
Label29.Caption = LanguageTx(96)
Label31.Caption = LanguageTx(97)
Label30.Caption = LanguageTx(98)
Command5.Caption = LanguageTx(99)
SaveBmpMnu.Caption = LanguageTx(100)
SetBackMnu.Caption = LanguageTx(101)
ExitSub.Caption = LanguageTx(102)
RenderMnu.Caption = LanguageTx(103)
RenderViewMnu.Caption = LanguageTx(104)
ClsImgMnu.Caption = LanguageTx(105)
EditMnu.Caption = LanguageTx(106)
LandEditorMnu.Caption = LanguageTx(107)
SkyEditorMnu.Caption = LanguageTx(108)
LightEditorMnu.Caption = LanguageTx(109)
MapEditorMnu.Caption = LanguageTx(110)
AboutMnu.Caption = LanguageTx(111)
LndMnu.Caption = LanguageTx(112)
Label26.Caption = LanguageTx(113)
End Sub

Private Sub ClsImgMnu_Click()
Call Command2_Click
End Sub

Private Sub Command1_Click()
On Error GoTo 1
SAV.DialogTitle = "Save as BMP"
SAV.ShowSave
SavePicture Image1.Picture, SAV.FileName + ".Bmp"
1:
End Sub

Private Sub Command2_Click()
On Error Resume Next
Image1.Picture = Me.Picture
Kill App.Path + "\Data\Textures\Render.Bmp"
End Sub

Private Sub Command3_Click()
Call RenderViewMnu_Click
End Sub

Private Sub Command4_Click()
SystemParametersInfo SPI_SETDESKWALLPAPER, 0, App.Path + "\Data\Textures\Render.BMP", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
End Sub

Private Sub Command5_Click()
Open App.Path + "\Data\Options\CameraOptions.Cnf" For Output As #1
Write #1, CSng(Text1.Text)
Write #1, CSng(Text2.Text)
Write #1, CSng(Text3.Text)
Write #1, CSng(Text5.Text)
Write #1, CSng(Text4.Text)
Close #1
End Sub

Private Sub ExitSub_Click()
End
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

Image1.Picture = LoadPicture(App.Path + "\Data\Textures\Render.Bmp")

Picture4.Picture = Picture8.Picture
Picture5.Picture = Picture8.Picture
Picture6.Picture = Picture8.Picture
Picture7.Picture = Picture8.Picture

Picture3.Picture = Picture2.Picture
Picture9.Picture = Picture2.Picture

Call Command5_Click

Combo1.ListIndex = 3
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label10_Click()
On Error GoTo 1
SAV.DialogTitle = "Open from CNF"
SAV.ShowOpen
FileCopy SAV.FileName, App.Path + "\Data\Options\LandOptions.CNF"
FileCopy Left(SAV.FileName, Len(SAV.FileName) - 4) + "2.Cnf", App.Path + "\Data\Options\LandOptions2.CNF"
1:
End Sub

Private Sub Label11_Click()
On Error GoTo 1
SAV.DialogTitle = "Open from CNF"
SAV.ShowOpen
FileCopy SAV.FileName, App.Path + "\Data\Options\LightOptions.CNF"
1:
End Sub

Private Sub Label12_Click()
On Error GoTo 1
SAV.DialogTitle = "Save as CNF"
SAV.ShowSave
FileCopy App.Path + "\Data\Options\LightOptions.CNF", SAV.FileName + ".Cnf"
1:
End Sub

Private Sub Label13_Click()
SetLightOpt.Show
End Sub

Private Sub Label16_Click()
SetSkyOpt.Show
End Sub

Private Sub Label17_Click()
On Error GoTo 1
SAV.DialogTitle = "Save as CNF"
SAV.ShowSave
FileCopy App.Path + "\Data\Options\SkyOptions.CNF", SAV.FileName + ".Cnf"
1:
End Sub

Private Sub Label18_Click()
On Error GoTo 1
SAV.DialogTitle = "Open from CNF"
SAV.ShowOpen
FileCopy SAV.FileName, App.Path + "\Data\Options\SkyOptions.CNF"
1:
End Sub

Private Sub Label19_Click()
On Error GoTo 1
SAV.DialogTitle = "Load Bmp Height Map"
SAV.ShowOpen
FileCopy SAV.FileName, App.Path + "\Data\Textures\HeightMap.BMP"
FileCopy Left(SAV.FileName, Len(SAV.FileName) - 4) + ".Lnd", App.Path + "\Data\Textures\LandMap.Lnd"
1:
End Sub

Private Sub Label20_Click()
On Error GoTo 1
SAV.DialogTitle = "Save Height Map"
SAV.ShowSave
FileCopy App.Path + "\Data\Textures\HeightMap.Bmp", SAV.FileName + ".Bmp"
FileCopy App.Path + "\Data\Textures\LandMap.Lnd", SAV.FileName + ".Lnd"
1:
End Sub

Private Sub Label21_Click()
SetHeightMap.Show
End Sub

Private Sub Label8_Click()
SetLandOpt.Show
End Sub

Private Sub Label9_Click()
On Error GoTo 1
SAV.DialogTitle = "Save as CNF"
SAV.ShowSave
FileCopy App.Path + "\Data\Options\LandOptions.CNF", SAV.FileName + ".Cnf"
FileCopy App.Path + "\Data\Options\LandOptions2.CNF", SAV.FileName + "2.Cnf"
1:
End Sub

Private Sub LandEditorMnu_Click()
SetLandOpt.Show
End Sub

Private Sub LightEditorMnu_Click()
SetLightOpt.Show
End Sub

Private Sub LndMnu_Click()
Splash.Show
Splash.Line4.Visible = False
Splash.Line5.Visible = False
Splash.Label3.Visible = False
Splash.Label10.Visible = True
Splash.Label8.Visible = True
Splash.Label7.Visible = True
Splash.Label6.Visible = True
Splash.Label5.Visible = True
Splash.Label4.Visible = True
End Sub

Private Sub MapEditorMnu_Click()
Call Label21_Click
End Sub

Private Sub NewWorldMnu_Click()
SetHeightMap.Show 1
End Sub

Private Sub OpenWorldMnu_Click()

End Sub

Private Sub RenderViewMnu_Click()
On Error GoTo 1
RenderLand.Show
1:
End Sub

Private Sub SaveWorldMnu_Click()

End Sub

Private Sub SaveBmpMnu_Click()
Call Command1_Click
End Sub

Private Sub SetBackMnu_Click()
Call Command4_Click
End Sub

Private Sub SkyEditorMnu_Click()
SetSkyOpt.Show
End Sub

