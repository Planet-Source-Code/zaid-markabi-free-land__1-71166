VERSION 5.00
Begin VB.Form SetHeightMap 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12735
   LinkTopic       =   "Form2"
   ScaleHeight     =   745
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   849
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check4 
      Appearance      =   0  'Flat
      Caption         =   "Auto Map"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      TabIndex        =   17
      Top             =   8400
      Width           =   1695
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      ScaleHeight     =   255
      ScaleWidth      =   1695
      TabIndex        =   15
      Top             =   8760
      Width           =   1695
      Begin VB.PictureBox Picture8 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1215
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load"
      Height          =   375
      Left            =   10920
      TabIndex        =   14
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10920
      TabIndex        =   13
      Top             =   10680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   10920
      TabIndex        =   12
      Top             =   9120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   10920
      TabIndex        =   11
      Top             =   10080
      Width           =   1695
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      Caption         =   "Auto Refresh"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      TabIndex        =   10
      Top             =   8040
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      Caption         =   "Land While Draw"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      TabIndex        =   9
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Low While Draw"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      TabIndex        =   8
      Top             =   7320
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   10800
      Picture         =   "SetHeightMap.frx":0000
      ScaleHeight     =   1785
      ScaleWidth      =   1905
      TabIndex        =   7
      Top             =   5400
      Width           =   1935
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   10800
      Picture         =   "SetHeightMap.frx":BC32
      ScaleHeight     =   1785
      ScaleWidth      =   1905
      TabIndex        =   6
      Top             =   3600
      Width           =   1935
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   10800
      Picture         =   "SetHeightMap.frx":15834
      ScaleHeight     =   1785
      ScaleWidth      =   1905
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   10800
      Picture         =   "SetHeightMap.frx":1C5F6
      ScaleHeight     =   1785
      ScaleWidth      =   1905
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10575
      Left            =   120
      ScaleHeight     =   705
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   705
      TabIndex        =   3
      Top             =   480
      Width           =   10575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   0
      Picture         =   "SetHeightMap.frx":22C38
      ScaleHeight     =   345
      ScaleWidth      =   10815
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   10440
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Height Map"
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
   Begin VB.Line Line5 
      BorderStyle     =   3  'Dot
      X1              =   720
      X2              =   800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderStyle     =   3  'Dot
      X1              =   848
      X2              =   848
      Y1              =   744
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      X1              =   720
      X2              =   720
      Y1              =   744
      Y2              =   16
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   848
      Y1              =   744
      Y2              =   744
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   0
      Y1              =   16
      Y2              =   744
   End
End
Attribute VB_Name = "SetHeightMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DrawEnable As Boolean
Dim HeightTerrain(703, 703) As Single
Dim TerrDraw(129, 129) As Single
Dim WeightSel As Integer




Dim picht As Integer
Dim picwt As Integer
Dim clflag As Boolean

Dim col1 As Long
Dim col2 As Long
Dim col3 As Long
Dim col As Integer

Dim XXX As Single
Dim YYY As Single


Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type

Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type
Dim sFile As String
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Sub LoadNewDoc()
Picture1.Picture = LoadPicture(sFile)
End Sub
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim r As Long
   Dim pic As PicBmp
   Dim IPic As IPicture
   Dim IID_IDispatch As GUID
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
   With pic
      .Size = Len(pic)
      .Type = vbPicTypeBitmap
      .hBmp = hBmp
      .hPal = hPal
   End With
   r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
   Set CreateBitmapPicture = IPic
End Function
  Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
  Dim hDCMemory As Long
  Dim hBmp As Long
  Dim hBmpPrev As Long
  Dim r As Long
  Dim hDCSrc As Long
  Dim hPal As Long
  Dim hPalPrev As Long
  Dim RasterCapsScrn As Long
  Dim HasPaletteScrn As Long
  Dim PaletteSizeScrn As Long
  Dim LogPal As LOGPALETTE
   If Client Then
      hDCSrc = GetDC(hWndSrc)
   Else
      hDCSrc = GetWindowDC(hWndSrc)
   End If
   hDCMemory = CreateCompatibleDC(hDCSrc)
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)
   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)                                                   ' capabilities.
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
   End If
   r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
   hBmp = SelectObject(hDCMemory, hBmpPrev)
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

Public Function CaptureScreen() As Picture
  Dim hWndScreen As Long
   hWndScreen = GetDesktopWindow()
   Set CaptureScreen = CaptureWindow(hWndScreen, False, (Me.Left / 15) + (Picture2.Left), (Me.Top / 15) + (Picture2.Top), Picture2.Width, Picture2.Height)
End Function

Function GreyScale(LongCol As Long) As Single
  Dim r As Single
  Dim g As Single
  Dim b As Single
  Long2RGB LongCol, r, g, b
  GreyScale = (r + b + g) / 765
End Function

Sub Long2RGB(LongCol As Long, r As Single, g As Single, b As Single)
  r = LongCol And 255
  g = (LongCol And 65280) \ 256&
  b = (LongCol And 16711680) \ 65535
End Sub









Private Sub RefreshLang()
Label2.Caption = LanguageTx(11)
Check1.Caption = LanguageTx(20)
Check2.Caption = LanguageTx(21)
Check3.Caption = LanguageTx(22)
Command2.Caption = LanguageTx(18)
Command3.Caption = LanguageTx(19)
Command4.Caption = LanguageTx(23)
Command1.Caption = LanguageTx(24)
Check4.Caption = LanguageTx(86)

End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
Picture2.Cls
Picture2.Picture = LoadPicture(App.Path + "\Data\Textures\HeightMap.JPG")
Open App.Path + "\Data\Textures\AutoMap.Lnd" For Output As #1
Write #1, "Y"
Close #1
Else
Call Command1_Click
Open App.Path + "\Data\Textures\AutoMap.Lnd" For Output As #1
Write #1, "N"
Close #1
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Picture8.Visible = True
DoEvents
Dim I1 As Integer
Dim I2 As Integer
For I1 = 0 To Picture2.ScaleWidth - 2
For I2 = 0 To Picture2.ScaleHeight - 2
If HeightTerrain(I1, I2) < 0 Then
HeightTerrain(I1, I2) = 0
End If
Picture2.ForeColor = RGB(HeightTerrain(I1, I2), HeightTerrain(I1, I2), HeightTerrain(I1, I2))
Picture2.PSet (I1, I2)
Next
Next
Picture8.Visible = False
Set Picture2.Picture = CaptureScreen()
End Sub

Private Sub Command2_Click()
SavePicture Picture2.Picture, App.Path + "\Data\Textures\HeightMap.Bmp"
Open App.Path + "\Data\Textures\LandMap.Lnd" For Output As #1
For I = 0 To 703
For I2 = 0 To 703
Write #1, HeightTerrain(I, I2)
Next
Next
Close #1

Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
Picture2.Picture = LoadPicture(App.Path + "\Data\Textures\HeightMap.Bmp")
Dim LndTm As Single
Open App.Path + "\Data\Textures\LandMap.Lnd" For Input As #1
For I = 0 To 703
For I2 = 0 To 703
Input #1, LndTm
HeightTerrain(I, I2) = LndTm
Next
Next
Close #1
End Sub

Private Sub Form_Load()
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

Call Picture4_Click
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Check1.Value = 1 Then
Picture2.DrawWidth = 6
End If
DrawEnable = True
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
If DrawEnable = True Then

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer

Dim ILndH As Integer
If Check2.Value = 1 Then
ILndH = 2
Else
ILndH = 1
End If

For I1 = 1 To WeightSel - 1
For I2 = 1 To WeightSel - 1

If Button = 1 Then
HeightTerrain(x + I1, Y + I2) = HeightTerrain(x + I1, Y + I2) + (TerrDraw(I1, I2) * 5)
Else
HeightTerrain(x + I1, Y + I2) = HeightTerrain(x + I1, Y + I2) - (TerrDraw(I1, I2) * 5)
End If

If Check1.Value = 0 Then
I3 = 11
End If

If I3 > 10 Then
Picture2.ForeColor = RGB(HeightTerrain(x + I1, Y + I2), HeightTerrain(x + I1, Y + I2), HeightTerrain(x + I1, Y + I2)) * ILndH
Picture2.PSet (Int(x) + I1, Int(Y) + I2)
I3 = 0
Else
I3 = I3 + 1
End If

Next
Next

End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Picture2.DrawWidth = 1

DrawEnable = False

If Check3 = 1 Then
Call Command1_Click
End If
End Sub

Private Sub Picture3_Click()
Dim I As Integer
Dim I2 As Integer
Dim x As Single

Open App.Path + "\Data\Draw Brush\SmallDraw.Tv8" For Input As #1
For I = 0 To 49
For I2 = 0 To 49
Input #1, x
TerrDraw(I, I2) = x
Next
Next
Close #1
WeightSel = 49

Picture3.Appearance = 1
Picture4.Appearance = 0
Picture5.Appearance = 0
Picture6.Appearance = 0
Picture3.BackColor = 0
Picture4.BackColor = 0
Picture5.BackColor = 0
Picture6.BackColor = 0
End Sub

Private Sub Picture4_Click()
Dim I As Integer
Dim I2 As Integer
Dim x As Single

Open App.Path + "\Data\Draw Brush\NormalDraw.Tv8" For Input As #1
For I = 0 To 65
For I2 = 0 To 65
Input #1, x
TerrDraw(I, I2) = x
Next
Next
Close #1
WeightSel = 65

Picture3.Appearance = 0
Picture4.Appearance = 1
Picture5.Appearance = 0
Picture6.Appearance = 0
Picture3.BackColor = 0
Picture4.BackColor = 0
Picture5.BackColor = 0
Picture6.BackColor = 0
End Sub

Private Sub Picture5_Click()
Dim I As Integer
Dim I2 As Integer
Dim x As Single

Open App.Path + "\Data\Draw Brush\LargeDraw.Tv8" For Input As #1
For I = 0 To 89
For I2 = 0 To 89
Input #1, x
TerrDraw(I, I2) = x
Next
Next
Close #1
WeightSel = 89

Picture3.Appearance = 0
Picture4.Appearance = 0
Picture5.Appearance = 1
Picture6.Appearance = 0
Picture3.BackColor = 0
Picture4.BackColor = 0
Picture5.BackColor = 0
Picture6.BackColor = 0
End Sub

Private Sub Picture6_Click()
Dim I As Integer
Dim I2 As Integer
Dim x As Single

Open App.Path + "\Data\Draw Brush\BigDraw.Tv8" For Input As #1
For I = 0 To 129
For I2 = 0 To 129
Input #1, x
TerrDraw(I, I2) = x
Next
Next
Close #1
WeightSel = 129

Picture3.Appearance = 0
Picture4.Appearance = 0
Picture5.Appearance = 0
Picture6.Appearance = 1
Picture3.BackColor = 0
Picture4.BackColor = 0
Picture5.BackColor = 0
Picture6.BackColor = 0
End Sub
