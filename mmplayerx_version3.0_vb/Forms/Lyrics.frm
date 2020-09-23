VERSION 5.00
Begin VB.Form frmLyrics 
   BorderStyle     =   0  'None
   Caption         =   "Lyrics"
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Lyrics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   295
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   30
      ScaleHeight     =   2145
      ScaleWidth      =   4365
      TabIndex        =   0
      Top             =   465
      Width           =   4365
      Begin VB.PictureBox picLyrics 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1890
         Left            =   0
         ScaleHeight     =   1890
         ScaleWidth      =   4320
         TabIndex        =   1
         Top             =   0
         Width           =   4320
         Begin VB.Shape shpFocus 
            BorderColor     =   &H00FF0000&
            Height          =   210
            Left            =   15
            Top             =   15
            Width           =   4305
         End
         Begin VB.Label lblLyrics 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "LyRiKs"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   30
            TabIndex        =   2
            Top             =   15
            Width           =   4270
         End
      End
      Begin VB.Label lblNoLyrics 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[ no lyrics found ]"
         Height          =   240
         Left            =   60
         TabIndex        =   3
         Top             =   945
         Visible         =   0   'False
         Width           =   4215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   30
      ScaleHeight     =   375
      ScaleWidth      =   4230
      TabIndex        =   4
      Top             =   15
      Width           =   4230
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Guns and Roses"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         TabIndex        =   6
         Top             =   165
         Width           =   4215
      End
      Begin VB.Label lblArtist 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Guns and Roses"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   -15
         TabIndex        =   5
         Top             =   0
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmLyrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Dim iLinesLyrics As Integer
Dim iCurrentLine As Integer

Dim lblForeColor As Long
Dim iLinesMax As Integer

Dim cWindows As New cWindowSkin
Dim cAjustarDesk As New clsDockingHandler
Dim InFormDrag As Boolean


Public Sub Reset_Values()
On Error Resume Next
 iLinesMax = 0
 '// Poner al estado normal
 If iCurrentLine > 0 Then lblLyrics(iCurrentLine).Font.Bold = False
 
 iCurrentLine = 0
 iLinesLyrics = frmMain.LyricsRef.ListCount
 picLyrics.Top = 0
 lblLyrics(0).Font.Bold = True
 shpFocus.Top = lblLyrics(0).Top
End Sub

Public Sub Move_Previous_Focus_Lyrics()
    
  iCurrentLine = iCurrentLine - 1
   If iCurrentLine < 0 Then
      iCurrentLine = 0
      iLinesMax = 0
      Exit Sub
   End If
 
   iLinesMax = iLinesMax - 1
   
   If iLinesMax < 0 Then
     iLinesMax = 9
     picLyrics.Top = (lblLyrics(0).Height * 10) - (lblLyrics(0).Height * (iCurrentLine + 1))
   End If
   
   shpFocus.Top = lblLyrics(iCurrentLine).Top
 
   lblLyrics(iCurrentLine + 1).Font.Bold = False
   lblLyrics(iCurrentLine).Font.Bold = True
End Sub


Public Sub Move_Next_Focus_Lyrics()
  iCurrentLine = iCurrentLine + 1
   If iCurrentLine > iLinesLyrics Then
      iCurrentLine = iLinesLyrics
      Exit Sub
   End If
 
   iLinesMax = iLinesMax + 1
   
   If iLinesMax > 9 Then
     iLinesMax = 0
     picLyrics.Top = -(lblLyrics(0).Height * iCurrentLine)
   End If
     
   shpFocus.Top = lblLyrics(iCurrentLine).Top
      
   lblLyrics(iCurrentLine - 1).Font.Bold = False
   lblLyrics(iCurrentLine).Font.Bold = True
End Sub

Public Sub Order_lblLyrics()
 Dim i As Integer
 Dim intHeight As Integer
 Dim strLyrics As String
 
 If frmMain.LyricsRef.ListCount = 0 Then Exit Sub
  lblForeColor = lblNoLyrics.ForeColor
  
   iLinesMax = 0
   iCurrentLine = 0
   iLinesLyrics = frmMain.LyricsRef.ListCount
   picLyrics.Top = 0
   shpFocus.Top = lblLyrics(0).Top
   lblLyrics(0).Font.Bold = False

  '// Todas los renglones de las letras
  For i = 0 To frmMain.LyricsRef.ListCount - 1
    If i >= lblLyrics.count Then
      '// Cargar nuevo label si todavia no esta
      Load lblLyrics(i)
    End If
    
    If i > 0 Then lblLyrics(i).Top = lblLyrics(i - 1).Top + lblLyrics(i - 1).Height
    strLyrics = Right(frmMain.LyricsRef.List(i), Len(frmMain.LyricsRef.List(i)) - 9)
    '// Configurar apariencia
    lblLyrics(i).Caption = strLyrics
    lblLyrics(i).ForeColor = lblForeColor
    lblLyrics(i).Visible = True
    intHeight = intHeight + lblLyrics(i).Height
  Next i
 picLyrics.Height = intHeight + 20
 lblLyrics(0).Font.Bold = True
End Sub


Private Sub Form_Load()
On Error Resume Next
 

  Me.Caption = LineLanguage(37)
  Me.Icon = frmMain.Icon
  frmLyrics.Left = (Screen.Width - frmLyrics.Width) / 2
  frmLyrics.Top = (Screen.Height - frmLyrics.Height) / 2

     
  Load_config_KARAOKE
  
  Set cWindows.FormularioPadre = Me
  Set cAjustarDesk.ParentForm = Me
  cargar_formulario
End Sub

Sub cargar_formulario()
On Error Resume Next
Dim iX As Integer, iY As Integer
  cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
  iX = Read_INI("Configuration", "ExitButtonX", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\KARAOKE\config.ini")
  iY = Read_INI("Configuration", "ExitButtonY", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\KARAOKE\config.ini")
  cWindows.ButtonExitXY CLng(iX), CLng(iY)
  cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\KARAOKE\config.ini")
  cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\KARAOKE\config.ini")
 cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\KARAOKE\"
   Picture1.Top = cWindows.AreaTop
   Picture1.Left = cWindows.AreaLeft
   Picture1.Width = cWindows.AreaWidth
   lblArtist.Width = Picture1.ScaleWidth
   lblTitle.Width = Picture1.ScaleWidth
   picBody.Top = cWindows.AreaTop + 25
   picBody.Left = cWindows.AreaLeft
   picBody.Width = cWindows.AreaWidth
   picLyrics.Width = picBody.ScaleWidth
   lblNoLyrics.Width = picBody.ScaleWidth
   Picture1.Height = cWindows.AreaHeight

End Sub


Sub Load_config_KARAOKE()
 On Error Resume Next
   
   bolLyricsShow = True
   
   picLyrics.BackColor = frmPlayList.PL.BackColor
   picBody.BackColor = picLyrics.BackColor

   shpFocus.BorderColor = frmPlayList.PL.ForeColor
   lblNoLyrics.ForeColor = shpFocus.BorderColor
 
 frmMain.LyricsIndex = 1
 
 If frmMain.LyricsRef.ListCount > 0 Then
      lblArtist.Caption = tCurrentID3.Artist & " - " & tCurrentID3.Album
      lblTitle.Caption = tCurrentID3.Title
      Order_lblLyrics
      picLyrics.Visible = True
      lblNoLyrics.Visible = False
 Else
      lblArtist.Caption = tCurrentID3.Artist & " - " & tCurrentID3.Album
      lblTitle.Caption = tCurrentID3.Title
      picLyrics.Visible = False
      lblNoLyrics.Visible = True
      lblNoLyrics.Caption = LineLanguage(38)
 End If
 
 frmLyrics.Visible = True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cWindows.Formulario_Down X, Y
cAjustarDesk.StartDockDrag X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
InFormDrag = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Y >= cWindows.AreaHeight + cWindows.AreaTop Then
    Exit Sub
 End If
 
 cWindows.Formulario_MouseMove Button, X, Y

 If cWindows.Ajustando = True Then
   Picture1.Width = cWindows.AreaWidth
   lblArtist.Width = Picture1.ScaleWidth
   lblTitle.Width = Picture1.ScaleWidth
   picBody.Width = cWindows.AreaWidth
   picLyrics.Width = picBody.ScaleWidth
   lblNoLyrics.Width = picBody.ScaleWidth

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
  Unload Me
End If
InFormDrag = False

End Sub

Private Sub Form_Paint()
cWindows.Formulario_Paint

End Sub

Private Sub Form_Unload(Cancel As Integer)
 bolLyricsShow = Not bolLyricsShow
 Me.Visible = bolLyricsShow
 frmPopUp.mnuLyrics.Checked = bolLyricsShow
 Cancel = 1
End Sub

Private Sub picLyrics_Resize()
 On Error Resume Next
 Dim i As Integer
 
 shpFocus.Width = picBody.ScaleWidth - 30
 For i = 0 To lblLyrics.count - 1
   lblLyrics(i).Width = shpFocus.Width - 25
 Next i
End Sub
