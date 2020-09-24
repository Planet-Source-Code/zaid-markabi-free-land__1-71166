VERSION 5.00
Begin VB.Form RenderLand 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   4545
   ClientTop       =   1845
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tree : -1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "RenderLand"
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

Public MS5 As Meshes_Sketch_English

Private Sub Form_Click()
MS5.Caputer_Screen App.Path + "\Data\Textures\Render.Bmp"
Main.Image1.Picture = LoadPicture(App.Path + "\Data\Textures\Render.Bmp")
MS5.Quit
Unload Me
End Sub

Private Sub Form_Load()
If Not Main.Combo1.ListIndex = 3 Then
Dim XX() As String
XX = Split(Main.Combo1.Text, "*")
Me.Width = Int(XX(0)) * 15
Me.Height = Int(XX(1)) * 15
Else
Me.Width = Screen.Width
Me.Height = Screen.Height
End If

    Me.Show
    
    Set MS5 = New Meshes_Sketch_English

    DoEvents
    
    MS5.Start Me.hwnd
    
   MS5.Effect_Fade_In 3
     
    MS5.Camera_Set_Position 1500, 0, 1500
     
    Open App.Path + "\Data\Options\LandOptions.Cnf" For Input As #1
Dim LandHeight As Single
     Input #1, x
     LandHeight = Int(x)
     Input #1, x2
Dim TextureFile As String
    Select Case x2
     Case Is = "0"
     TextureFile = App.Path + "\Data\Textures\Land.JPG"
     Case Is = "1"
     TextureFile = App.Path + "\Data\Textures\Sand.JPG"
     Case Is = "2"
     TextureFile = App.Path + "\Data\Textures\Grass.JPG"
     Case Is = "3"
     TextureFile = App.Path + "\Data\Textures\Snow.JPG"
     Case Is = "4"
     TextureFile = App.Path + "\Data\Textures\Natural.JPG"
     Case Is = "5"
     TextureFile = App.Path + "\Data\Textures\Gold.JPG"
     Case Is = "6"
     TextureFile = App.Path + "\Data\Textures\Red.JPG"
     Case Is = "7"
     TextureFile = "1"
     Case Is = "8"
     TextureFile = "2"
     Case Is = "9"
     TextureFile = "3"
     Case Is = "10"
     TextureFile = "4"
    End Select

     Input #1, x2
     
Dim LandWidth As Integer
     
Input #1, x
LandWidth = Int(x)
Input #1, x4
     Input #1, x4
Dim LandQutly As Integer
Dim LandSmooth As Boolean
      If x4 = "0" Then
    LandQutly = 4
    LandSmooth = False
      End If
      If x4 = "1" Then
    LandQutly = 2
    LandSmooth = False
      End If
      If x4 = "2" Then
    LandQutly = 1
    LandSmooth = False
      End If
      If x4 = "3" Then
    LandQutly = 4
    LandSmooth = True
      End If
      If x4 = "4" Then
    LandQutly = 0
    LandSmooth = False
      End If
      If x4 = "5" Then
    LandQutly = 3
    LandSmooth = False
      End If

     Close #1

Open App.Path + "\Data\Options\LightOptions.Cnf" For Input As #1
     Input #1, x
If x = "1" Then

Dim LightPower As Single
     Input #1, x
     LightPower = CSng(x)

Dim LightColorR As Single
Dim LightColorG As Single
Dim LightColorB As Single
     Input #1, x
     LightColorR = CSng(x)
     Input #1, x
     LightColorG = CSng(x)
     Input #1, x
     LightColorB = CSng(x)
     Input #1, x
Dim ShadowPower As Single
     ShadowPower = CSng(x)
     Input #1, x
 Select Case x
Dim LightDirection As String
  Case Is = "0"
    LightDirection = "U"
  Case Is = "1"
    LightDirection = "D"
  Case Is = "2"
    LightDirection = "R"
  Case Is = "3"
    LightDirection = "L"
  Case Is = "4"
    LightDirection = "A"
 End Select
Dim LightRot As Single
     Input #1, x
LightRot = CSng(x)
  End If
Close #1

    Open App.Path + "\Data\Options\SkyOptions.Cnf" For Input As #1
Dim SkyEnbl As Boolean
     Input #1, x
     SkyEnbl = x
Dim SkyQuitly As Integer
     Input #1, x
     SkyQuitly = Int(x)
     Input #1, x
     If x = "1" Then
Dim SkyRotate As Single
     Input #1, x
     SkyRotate = CSng(x)
     Else
     Input #1, x
     SkyRotate = 0
     End If
     Input #1, x
     Dim SkyModeColor As Integer
     SkyModeColor = x
    If x = "0" Then
      MS5.World_Set_BackGround 40, 60, 90
    End If
    If x = "1" Then
      MS5.World_Set_BackGround 60, 70, 90
    End If
    If x = "2" Then
      MS5.World_Set_BackGround 20, 30, 45
    End If
    If x = "3" Then
      MS5.World_Set_BackGround 10, 15, 22
    End If
    Input #1, x
    Dim SphereSky As Boolean
    SphereSky = x
    Input #1, x
    Dim CloudeLevel As Single
    CloudeLevel = x
    Input #1, x
    Dim PowerSky1 As Single
    PowerSky1 = x
    Input #1, x
    Dim PowerSky2 As Single
    PowerSky2 = x
    Close #1


Open App.Path + "\Data\Options\LandOptions2.Cnf" For Input As #1
     Input #1, x
    If x = "0" Then
Dim DetailFile As String
DetailFile = App.Path + "\Data\Textures\Detail High.Jpg"
    Else
DetailFile = App.Path + "\Data\Textures\Detail Low.Jpg"
    End If
Dim TextureSize As Single
     Input #1, x
     TextureSize = CSng(x)
     Input #1, x
Dim DetailSize As Single
     Input #1, x
     DetailSize = CSng(x)
     Input #1, x
     Input #1, x
     Select Case x
     Case Is = "0"
Dim DetailMode As Integer
DetailMode = 0
     Case Is = "1"
DetailMode = 1
     Case Is = "2"
DetailMode = 2
     Case Is = "3"
DetailMode = 3
     End Select
Input #1, x
Dim TreeOn As Boolean
TreeOn = x
Input #1, x
Dim TreeDist As Integer
TreeDist = x
Close #1

Open App.Path + "\Data\Textures\AutoMap.Lnd" For Input As #1
Input #1, x
Dim MapFile As String
If x = "Y" Then
MapFile = App.Path + "\Data\Textures\HeightMap.JPG"
    MS5.Land_Create False, TextureFile, TextureSize, App.Path + "\Data\Textures\HeightMap.JPG", LandQutly, False, LandWidth, LandHeight
Else
MapFile = App.Path + "\Data\Textures\HeightMap.BMP"
    MS5.Land_Create False, TextureFile, TextureSize, App.Path + "\Data\Textures\HeightMap.BMP", LandQutly, False, LandWidth, LandHeight
End If
Close #1
    
    MS5.Land_Add_Detail DetailFile, DetailSize, DetailMode
    
    MS5.Land_Add_3D_Light LightPower, LightColorR * 100, LightColorG * 100, LightColorB * 100, ShadowPower, LightDirection

    MS5.Land_Add_Water 10, 50, True, App.Path + "\Data\Textures\water" + Format(SkyModeColor) + ".jpg", 20, 0, 1000
     
If SkyEnbl = True Then
     If SphereSky = True Then
    MS5.World_Add_Sky_Sphere App.Path + "\Data\Textures\Cloud.dds", SkyQuitly, SkyRotate
     Else
    MS5.World_Create_Sky_Texture App.Path + "\Data\Textures\Sky" + Format(SkyModeColor) + "1.bmp", App.Path + "\Data\Textures\Sky" + Format(SkyModeColor) + "2.bmp", CloudeLevel, PowerSky1, PowerSky2, App.Path + "\Data\Textures\Sky Map.jpg", "C:\TF.Bmp"
    MS5.World_Add_Sky "C:\TF.Bmp", "C:\TF.Bmp", "C:\TF.Bmp", "C:\TF.Bmp", "C:\TF.Bmp", "C:\TF.Bmp"
     End If
End If
     
If TreeOn = True Then
Me.Show
DoEvents
Me.Label1.Visible = True
Dim TreeNum As Integer
MS5.World_Load_Texture App.Path + "\Data\Tree\Tree.Bmp", "Tree", True
For I = 0 To 2500
For I2 = 0 To 2500
TreeNum = TreeNum + 1
MS5.Object_Add "Tree" + Format(TreeNum)
MS5.Object_Open_From_File "Tree" + Format(TreeNum), App.Path + "\Data\Tree\Tree.3DS", False, "", False, False
MS5.Object_Set_Texture "Tree" + Format(TreeNum), "Tree", -1
MS5.Object_Set_Position "Tree" + Format(TreeNum), CSng(I), MS5.Land_Get_Height(CSng(I), CSng(I2)), CSng(I2)
MS5.Object_Set_Rotation "Tree" + Format(TreeNum), 0, Rnd * 1000, 0
I2 = I2 + Int(Rnd * TreeDist)
Me.Label1.Caption = "Tree : " + Format(TreeNum)
DoEvents
Next
I = I + Int(Rnd * TreeDist)
Next
Me.Label1.Visible = False
End If

Select Case TextureFile
Case Is = "1"
MS5.Land_Set_Layer_2 App.Path + "\Data\Textures\Grass.JPG", App.Path + "\Data\Textures\Land.JPG", MapFile, 8, 70, 0.8, 0.6
Case Is = "2"
MS5.Land_Set_Layer_2 App.Path + "\Data\Textures\Red.JPG", App.Path + "\Data\Textures\Snow.JPG", MapFile, 8, 135, 0.7, 0.7
Case Is = "3"
MS5.Land_Set_Layer_2 App.Path + "\Data\Textures\Grass.JPG", App.Path + "\Data\Textures\Natural.JPG", MapFile, 7, 70, 0.8, 0.6
Case Is = "4"
MS5.Land_Set_Layer_2 App.Path + "\Data\Textures\Gold.jpg", App.Path + "\Data\Textures\Red.JPG", MapFile, 8, 135, 0.5, 1

End Select


Dim CamSpeed As Single
Dim CamHeightEye As Single
Dim CamCrackling As Single
Dim CamCorner As Single
Dim CamFarV As Single
Open App.Path + "\Data\Options\CameraOptions.Cnf" For Input As #1
Input #1, CamSpeed
Input #1, CamHeightEye
Input #1, CamCrackling
Input #1, CamCorner
Input #1, CamFarV
Close #1
MS5.Camera_Set_View CamCorner, CamFarV

     Do While MS5.Is_Working_Now = True
       MS5.Render
     
     MS5.Walk_In_The_World_Freely False, CamSpeed, True, True, CamHeightEye, CamCrackling

If LightRot > 0 Then
     MS5.Light_Rotate 1, LightRot * 1000
End If
     
        DoEvents
     Loop

End Sub
