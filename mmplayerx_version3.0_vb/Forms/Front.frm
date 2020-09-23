VERSION 5.00
Begin VB.Form frmCaratula 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   " Cover Front"
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   4320
   FontTransparent =   0   'False
   Icon            =   "Front.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   288
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picfondo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   165
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   1
      Top             =   240
      Width           =   2865
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   2085
      ScaleHeight     =   123
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   140
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   2100
   End
End
Attribute VB_Name = "frmCaratula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cWindows As New cWindowSkin
Dim cAjustarDesk As New clsDockingHandler
Dim InFormDrag As Boolean

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cWindows.Formulario_Down X, Y
cAjustarDesk.StartDockDrag X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
InFormDrag = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 cWindows.Formulario_MouseMove Button, X, Y

 If cWindows.Ajustando = True Then
  picfondo.PaintPicture Picture1.Picture, 0, 0, cWindows.AreaWidth, cWindows.AreaHeight, 0, 0

    picfondo.Left = cWindows.AreaLeft
    picfondo.Top = cWindows.AreaTop
    picfondo.Width = cWindows.AreaWidth
    picfondo.Height = cWindows.AreaHeight
 
 End If
    If InFormDrag And cWindows.Ajustando = False Then
        ' Continue window draggin'
        cAjustarDesk.UpdateDockDrag X * Screen.TwipsPerPixelX, _
            Y * Screen.TwipsPerPixelY
        Exit Sub
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cWindows.Formulario_MouseUp X, Y
If cWindows.ClickExitButton = True Then
 bolCaratulaShow = Not bolCaratulaShow
 frmCaratula.Visible = bolCaratulaShow
 frmPopUp.mnuMCaratula.Checked = bolCaratulaShow
End If

InFormDrag = False
End Sub

Private Sub Form_Paint()
cWindows.Formulario_Paint
End Sub


Private Sub Form_Load()
On Error Resume Next
  
 '// si el album tiene caratula mostrarla
 If Trim(strRutaCaratula) <> "" Then
   Picture1.Picture = LoadPicture(strRutaCaratula)
 Else
   '// si no tiene caratula el album mostrar el default logo
   Picture1.Picture = frmPopUp.picDefaultLogo.Picture
 End If
   Picture1.AutoSize = True
'  Me.Width = Picture1.Width * 15: Me.Height = Picture1.Height * 15

  Me.Left = (Screen.Width - Me.Width) / 2   '// centrar formulario
  Me.Top = (Screen.Height - Me.Height) / 2
  

  Set cWindows.FormularioPadre = Me
  Set cAjustarDesk.ParentForm = Me
  cargar_formulario
  bolCaratulaShow = True

End Sub

Sub cargar_formulario()
On Error Resume Next
Dim iX As Integer, iY As Integer
  cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
  iX = Read_INI("Configuration", "ExitButtonX", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\COVER_FRONT\config.ini")
  iY = Read_INI("Configuration", "ExitButtonY", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\COVER_FRONT\config.ini")
  cWindows.ButtonExitXY CLng(iX), CLng(iY)
  cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\COVER_FRONT\config.ini")
  cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\COVER_FRONT\config.ini")
 cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\COVER_FRONT\"

    picfondo.Left = cWindows.AreaLeft
    picfondo.Top = cWindows.AreaTop
    picfondo.Width = cWindows.AreaWidth
    picfondo.Height = cWindows.AreaHeight

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Resize()
 Mover_Form
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Sub Mover_Form()
 '// ajustar la imagen al ancho alto del form
 picfondo.PaintPicture Picture1.Picture, 0, 0, picfondo.ScaleWidth, picfondo.ScaleHeight, 0, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
 bolCaratulaShow = False
 Set cWindows = Nothing
 Set cAjustarDesk = Nothing

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'Private Sub picfondo_DblClick()
' '// ajustar el formulario al ancho-alto original de la caratula
'   Picture1.AutoSize = True
'   Me.Width = Picture1.Width * Screen.TwipsPerPixelX
'   Me.Height = Picture1.Height * Screen.TwipsPerPixelY
'
'End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
