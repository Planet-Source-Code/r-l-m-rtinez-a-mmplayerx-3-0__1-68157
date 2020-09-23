VERSION 5.00
Begin VB.Form frmSpectrum 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Spectrum"
   ClientHeight    =   2010
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSpectrum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   0
      ScaleHeight     =   117
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   266
      TabIndex        =   1
      Top             =   0
      Width           =   3990
      Begin VB.Label Label 
         BackColor       =   &H00000000&
         Caption         =   "Visualizacion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   4245
      End
   End
   Begin VB.PictureBox picFront 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   2895
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   0
      Top             =   2835
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4860
      Top             =   1080
   End
End
Attribute VB_Name = "frmSpectrum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bLoadingVis As Boolean
Dim InFormDrag As Boolean
Dim cWindows As New cWindowSkin
Dim cAjustarDesk As New clsDockingHandler

Private Sub Form_Resize()
 On Error Resume Next

' If tConfigVis.Exist = False Then Exit Sub

' picSpectrum.Cls
' picSpectrum.PaintPicture picFront.Picture, 0, 0, picSpectrum.ScaleWidth, picSpectrum.ScaleHeight, 0, 0
' picSpectrum.Picture = picSpectrum.Image

 DoEvents

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cWindows.Formulario_Down X, Y
            cAjustarDesk.StartDockDrag X * Screen.TwipsPerPixelX, _
                Y * Screen.TwipsPerPixelY
InFormDrag = True
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cWindows.Formulario_MouseMove Button, X, Y

 If cWindows.Ajustando = True Then
  If tConfigVis.Exist = True Then
   picSpectrum.Cls
   picSpectrum.PaintPicture picFront.Picture, 0, 0, cWindows.AreaWidth, cWindows.AreaHeight, 0, 0
   picSpectrum.Picture = picSpectrum.Image
  End If
    picSpectrum.Left = cWindows.AreaLeft
    picSpectrum.Top = cWindows.AreaTop
    picSpectrum.Width = cWindows.AreaWidth
    picSpectrum.Height = cWindows.AreaHeight
    Label.Width = cWindows.AreaWidth

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
frmMain.Mostrar_Visualizacion

End If
InFormDrag = False
    
End Sub


Private Sub picSpectrum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu frmPopUp.mnuSpectrum

End Sub

Private Sub picSpectrum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbDefault
End Sub


Private Sub Form_Paint()
cWindows.Formulario_Paint
    

End Sub

Sub Load_Visualizacion(sFileVis As String)
 On Error Resume Next
 Dim s As String, i As Integer
 Dim bExistConfigScope As Boolean
 
 bLoadingVis = True
 
 sFileVis = tAppConfig.AppConfig & "Settings\" & sFileVis & ".vis"
 
 If Dir(sFileVis) = "" Then Exit Sub
 
 i = 0
 ' OSCILLOSCOPE
 ReDim tConfigScope(-1)
 Do
    s = Read_INI("Oscilloscope_" & i, "Number", "", , , sFileVis)
    If s <> "" Then
      ReDim Preserve tConfigScope(i)
       tConfigScope(i).Align = Read_INI("Oscilloscope_" & i, "Align", 1, , , sFileVis)
       If tConfigScope(i).Align < 0 Or tConfigScope(i).Align > 2 Then tConfigScope(i).Align = 1
       tConfigScope(i).BackColorScope = Read_INI("Oscilloscope_" & i, "BackColorScope", RGB(0, 255, 0), , , sFileVis)
       tConfigScope(i).LinesScope = Read_INI("Oscilloscope_" & i, "LinesScope", 50, , , sFileVis)
       If tConfigScope(i).LinesScope < 6 Or tConfigScope(i).LinesScope > 200 Then tConfigScope(i).LinesScope = 50
       bExistConfigScope = True
    End If
    i = i + 1
  Loop While s <> ""
   
 ' SPECTRUM
 With tConfigVis
  .BackColor = Read_INI("Spectrum", "BackColor", RGB(0, 0, 0), , , sFileVis)
  .BackColorBar = Read_INI("Spectrum", "BackColorBar", RGB(255, 255, 255), , , sFileVis)
  .BackColorPeak = Read_INI("Spectrum", "BackColorPeak", RGB(255, 255, 255), , , sFileVis)
  .Bars = Read_INI("Spectrum", "Bars", 50, , , sFileVis)
  If .Bars < 6 Or .Bars > 200 Then .Bars = 50
  .DrawBars = CBool(Read_INI("Spectrum", "DrawBars", 1, , , sFileVis))
  .DrawPeaks = CBool(Read_INI("Spectrum", "DrawPeaks", 1, , , sFileVis))
  .DrawSource = Read_INI("Spectrum", "DrawSource", 1, , , sFileVis)
  .Exist = CBool(Read_INI("Spectrum", "Exist", 1, , , sFileVis))
  .Gradient = Read_INI("Spectrum", "Gradient", "No Hay.jpg", , , sFileVis)
  .GrandientIndex = Read_INI("Spectrum", "GradientIndex", 0, , , sFileVis)
  .ImageFile = Read_INI("Spectrum", "ImageFile", "[Cover Front]", , , sFileVis)
  .Mirrored = CBool(Read_INI("Spectrum", "Mirrored", 1, , , sFileVis))
  .PeakGravity = Read_INI("Spectrum", "PeakGravity", 2, , , sFileVis)
  If .PeakGravity < 0 Or .PeakGravity > 4 Then .PeakGravity = 3
  .PeakHeight = Read_INI("Spectrum", "PeakHeight", 1, , , sFileVis)
  If .PeakHeight < 0 Or .PeakHeight > 4 Then .PeakHeight = 2
  .ScaleUp = Read_INI("Spectrum", "ScaleUp", 0, , , sFileVis)
  .Spacio = Read_INI("Spectrum", "Space", 0, , , sFileVis)
  If .Spacio > 10 Then .Spacio = 10
  
  
  If .DrawBars = False And .DrawPeaks = False Then .DrawBars = True

  
  ReDim tConfigVis.arryPeaks(tConfigVis.Bars)
  ReDim tConfigVis.arryWaitPeak(tConfigVis.Bars)
  Setup_Visualizacion

  bLoadingVis = False
End With


End Sub
Sub Setup_Visualizacion()
 On Error Resume Next
   picSpectrum.Picture = LoadPicture()
   picFront.Picture = LoadPicture()
     
 If tConfigVis.Exist = True Then
   picSpectrum.BackColor = tConfigVis.BackColorBar
     
   If tConfigVis.DrawBars = True Then
     '// gradient
      If tConfigVis.DrawSource = 0 Then
          picFront.Picture = LoadPicture(tAppConfig.AppConfig & "Settings\" & tConfigVis.Gradient)
          '// Image
       ElseIf tConfigVis.DrawSource = 1 Then
              If Dir(tConfigVis.ImageFile) <> "" Then
                 picFront.Picture = LoadPicture(tConfigVis.ImageFile)
              Else
                 If Trim(strRutaCaratula) <> "" Then
                    picFront.Picture = LoadPicture(strRutaCaratula)
                 Else '// si no tiene caratula el album mostrar el default logo
                    picFront.Picture = frmPopUp.picDefaultLogo.Picture
                 End If
              End If
           End If
    Else
      picSpectrum.BackColor = tConfigVis.BackColor
    End If
   picFront.AutoSize = True
   picSpectrum.Cls
   picSpectrum.PaintPicture picFront.Picture, 0, 0, cWindows.AreaWidth, cWindows.AreaHeight, 0, 0
   picSpectrum.Picture = picSpectrum.Image
 Else
   picSpectrum.BackColor = 0
 End If
  

End Sub
Public Sub Stop_Visualizacion()
Dim X1 As Single, Y1 As Single
Dim X2 As Single, Y2 As Single
Dim i As Integer, iSleep As Integer, j As Integer
Dim iSpacio As Integer, iPeak As Single, RaiseBars As Single, RaiseBars2 As Single
Dim Max&

'On Error Resume Next
On Error GoTo hell
   
If bLoadingVis = True Then Exit Sub

picSpectrum.Cls
  
'// SPECTRUM ANALYZER
If tConfigVis.Exist = True Then
  For i = 0 To tConfigVis.Bars
      X1 = i * (picSpectrum.ScaleWidth / tConfigVis.Bars)
      X2 = X1 + (picSpectrum.ScaleWidth / tConfigVis.Bars)
      '---------------------------------------------------------------------
      '// full window
      If tConfigVis.Mirrored = True Then
         Y1 = picSpectrum.ScaleHeight / 2
      Else
         Y1 = picSpectrum.ScaleHeight
      End If
            
      '---------------------------------------------------------------------
      '// Raise bars
      If tConfigVis.ScaleUp = 0 Then 'Normal
         RaiseBars = Y1
         RaiseBars2 = Y1
      ElseIf tConfigVis.ScaleUp = 1 Then
             RaiseBars = (picSpectrum.ScaleHeight / 3)
             RaiseBars2 = Y1 + (picSpectrum.ScaleHeight / 5)
          ElseIf tConfigVis.ScaleUp = 2 Then
                 RaiseBars = (picSpectrum.ScaleHeight / 6)
                 RaiseBars2 = (Y1 + (picSpectrum.ScaleHeight / 6) * 2)
              ElseIf tConfigVis.ScaleUp = 3 Then
                     RaiseBars = (picSpectrum.ScaleHeight / 10)
                     RaiseBars2 = (Y1 + (picSpectrum.ScaleHeight / 10) * 4)
                  End If
      '---------------------------------------------------------------------
      Max = (0 * RaiseBars)
                        
                        
      If Max >= Y1 And tConfigVis.DrawPeaks = True Then Max = Y1 - tConfigVis.PeakHeight
                     
     '====================================================================
     '// bars
     If tConfigVis.DrawBars = True Then

        Y2 = RaiseBars - Max
        picSpectrum.Line (X1, Y2)-(X2, 0), tConfigVis.BackColor, BF
        
        If tConfigVis.Spacio >= 0 Then
          picSpectrum.Line (X2, Y1)-(X2 + tConfigVis.Spacio, 0), tConfigVis.BackColor, BF
        End If
       
       '// espejo
        If tConfigVis.Mirrored = True Then
           Y2 = RaiseBars2 + Max
           picSpectrum.Line (X1, Y2)-(X2, Y1 * 2), tConfigVis.BackColor, BF
           
           If tConfigVis.Spacio >= 0 Then
             picSpectrum.Line (X2, Y1)-(X2 + tConfigVis.Spacio, Y1 * 2), tConfigVis.BackColor, BF
           End If
        Else
           picSpectrum.Line (X1, Y1)-(X2, Y1 * 2), tConfigVis.BackColor, BF
        End If
     End If
          
     If tConfigVis.Spacio >= 0 Then
        X2 = X2 - 1
        X1 = X1 + 1 + tConfigVis.Spacio
     End If
     
     '====================================================================
     '// Peaks
     
     If tConfigVis.DrawPeaks = True Then
       tConfigVis.arryPeaks(i) = 0
       iPeak = RaiseBars - tConfigVis.arryPeaks(i)
       picSpectrum.Line (X1, iPeak - 1)-(X2, iPeak - tConfigVis.PeakHeight), tConfigVis.BackColorPeak, BF
       '// peaks de espejo
       If tConfigVis.Mirrored = True Then
          iPeak = RaiseBars2 + tConfigVis.arryPeaks(i) - tConfigVis.PeakHeight
          picSpectrum.Line (X1, iPeak + 1)-(X2, iPeak + tConfigVis.PeakHeight), tConfigVis.BackColorPeak, BF
       End If
    End If

  Next i
End If

'================================================================================
'// OSCILLOSCOPE
For j = 0 To UBound(tConfigScope)
  For i = 0 To tConfigScope(j).LinesScope
     X1 = i * (picSpectrum.ScaleWidth / tConfigScope(j).LinesScope)
     X2 = X1 + (picSpectrum.ScaleWidth / tConfigScope(j).LinesScope)
     Y1 = picSpectrum.ScaleHeight
      
     '// full window
     If tConfigScope(j).Align = 1 Then Y1 = picSpectrum.ScaleHeight / 2
      
     '// top y bottom bars
     If tConfigScope(j).Align = 0 Or tConfigScope(j).Align = 2 Then Y1 = picSpectrum.ScaleHeight / 4
      
     Y2 = 0
    
     '// bottom align
     If tConfigScope(j).Align = 0 Then Y1 = Y1 * 3
        

     picSpectrum.Line (X1, Y1)-(X1 + ((X2 - X1) / 3), Y1 - Y2), tConfigScope(j).BackColorScope
     picSpectrum.Line (X1 + ((X2 - X1) / 3), Y1 - Y2)-(X1 + (((X2 - X1) / 3) * 2), Y1 + Y2), tConfigScope(j).BackColorScope
     picSpectrum.Line (X1 + (((X2 - X1) / 3) * 2), Y1 + Y2)-(X2, Y1), tConfigScope(j).BackColorScope
  Next i
Next j
 
Exit Sub
hell:
End Sub
Public Sub Update_Visualizacion(arryValues() As Single)
Dim X1 As Single, Y1 As Single
Dim X2 As Single, Y2 As Single
Dim i As Integer, iSleep As Integer, j As Integer
Dim iSpacio As Integer, iPeak As Single, RaiseBars As Single, RaiseBars2 As Single
Dim Max&

'On Error Resume Next
On Error GoTo hell
If bLoadingVis = True Or cWindows.Ajustando = True Then Exit Sub
   
picSpectrum.Cls
  
'// SPECTRUM ANALYZER
If tConfigVis.Exist = True Then
  For i = 0 To tConfigVis.Bars
      X1 = i * (picSpectrum.ScaleWidth / tConfigVis.Bars)
      X2 = X1 + (picSpectrum.ScaleWidth / tConfigVis.Bars)
      '---------------------------------------------------------------------
      '// full window
      If tConfigVis.Mirrored = True Then
         Y1 = picSpectrum.ScaleHeight / 2
      Else
         Y1 = picSpectrum.ScaleHeight
      End If
            
      '---------------------------------------------------------------------
      '// Raise bars
      If tConfigVis.ScaleUp = 0 Then 'Normal
         RaiseBars = Y1
         RaiseBars2 = Y1
      ElseIf tConfigVis.ScaleUp = 1 Then
             RaiseBars = (picSpectrum.ScaleHeight / 3)
             RaiseBars2 = Y1 + (picSpectrum.ScaleHeight / 5)
          ElseIf tConfigVis.ScaleUp = 2 Then
                 RaiseBars = (picSpectrum.ScaleHeight / 6)
                 RaiseBars2 = (Y1 + (picSpectrum.ScaleHeight / 6) * 2)
              ElseIf tConfigVis.ScaleUp = 3 Then
                     RaiseBars = (picSpectrum.ScaleHeight / 10)
                     RaiseBars2 = (Y1 + (picSpectrum.ScaleHeight / 10) * 4)
                  End If
      '---------------------------------------------------------------------
      Max = (Format(arryValues(i), ".00") * RaiseBars)
                        
      'Max = Max * (tConfigVis.ScaleUp+1)
                        
      If Max >= Y1 And tConfigVis.DrawPeaks = True Then Max = Y1 - tConfigVis.PeakHeight
                     
     '====================================================================
     '// bars
     If tConfigVis.DrawBars = True Then

        Y2 = RaiseBars - Max
        picSpectrum.Line (X1, Y2)-(X2, 0), tConfigVis.BackColor, BF
        
        If tConfigVis.Spacio >= 0 Then
          picSpectrum.Line (X2, Y1)-(X2 + tConfigVis.Spacio, 0), tConfigVis.BackColor, BF
        End If
       
       '// espejo
        If tConfigVis.Mirrored = True Then
           Y2 = RaiseBars2 + Max
           picSpectrum.Line (X1, Y2)-(X2, Y1 * 2), tConfigVis.BackColor, BF
           
           If tConfigVis.Spacio >= 0 Then
             picSpectrum.Line (X2, Y1)-(X2 + tConfigVis.Spacio, Y1 * 2), tConfigVis.BackColor, BF
           End If
        Else
           picSpectrum.Line (X1, Y1)-(X2, Y1 * 2), tConfigVis.BackColor, BF
        End If
     End If
          
     If tConfigVis.Spacio >= 0 Then
        X2 = X2 - 1
        X1 = X1 + 1 + tConfigVis.Spacio
     End If
     
     '====================================================================
     '// Peaks
     
     If tConfigVis.DrawPeaks = True Then
       If tConfigVis.arryPeaks(i) < Max Then
          tConfigVis.arryPeaks(i) = Max
          tConfigVis.arryWaitPeak(i) = Time
       End If

       If tConfigVis.arryPeaks(i) < 0 Then tConfigVis.arryPeaks(i) = 0
        
       iPeak = RaiseBars - tConfigVis.arryPeaks(i)
     
       If iPeak <= tConfigVis.PeakHeight Then iPeak = tConfigVis.PeakHeight
            
       picSpectrum.Line (X1, iPeak - 1)-(X2, iPeak - tConfigVis.PeakHeight), tConfigVis.BackColorPeak, BF
     
       '// peaks de espejo
       If tConfigVis.Mirrored = True Then
          iPeak = RaiseBars2 + tConfigVis.arryPeaks(i) - tConfigVis.PeakHeight
          If iPeak >= picSpectrum.ScaleHeight Then iPeak = picSpectrum.ScaleHeight - tConfigVis.PeakHeight - 1
          
          picSpectrum.Line (X1, iPeak + 1)-(X2, iPeak + tConfigVis.PeakHeight), tConfigVis.BackColorPeak, BF
       End If
       
         If tConfigVis.arryWaitPeak(i) <> "" Then iSleep = DateDiff("s", tConfigVis.arryWaitPeak(i), Time)
         If (iSleep >= 1) Then tConfigVis.arryPeaks(i) = tConfigVis.arryPeaks(i) - tConfigVis.PeakGravity
     End If

  Next i
End If

'================================================================================
'// OSCILLOSCOPE
For j = 0 To UBound(tConfigScope)
  For i = 0 To tConfigScope(j).LinesScope
     X1 = i * (picSpectrum.ScaleWidth / tConfigScope(j).LinesScope)
     X2 = X1 + (picSpectrum.ScaleWidth / tConfigScope(j).LinesScope)
     Y1 = picSpectrum.ScaleHeight
      
     '// full window
     If tConfigScope(j).Align = 1 Then Y1 = picSpectrum.ScaleHeight / 2
      
     '// top y bottom bars
     If tConfigScope(j).Align = 0 Or tConfigScope(j).Align = 2 Then Y1 = picSpectrum.ScaleHeight / 4
      
     Y2 = (Format(arryValues(i), ".00") * Y1)
    
     '// bottom align
     If tConfigScope(j).Align = 0 Then Y1 = Y1 * 3
        

     picSpectrum.Line (X1, Y1)-(X1 + ((X2 - X1) / 3), Y1 - Y2), tConfigScope(j).BackColorScope
     picSpectrum.Line (X1 + ((X2 - X1) / 3), Y1 - Y2)-(X1 + (((X2 - X1) / 3) * 2), Y1 + Y2), tConfigScope(j).BackColorScope
     picSpectrum.Line (X1 + (((X2 - X1) / 3) * 2), Y1 + Y2)-(X2, Y1), tConfigScope(j).BackColorScope
  Next i

Next j

Exit Sub
hell:

End Sub

Private Sub Form_Load()
  Me.Left = (Screen.Width - Me.Width) / 2   '// centrar formulario
  Me.Top = (Screen.Height - Me.Height) / 2
  

End Sub
Sub Cargar_formulario()
On Error Resume Next
Dim iX As Integer, iY As Integer

  Set cWindows.FormularioPadre = Me
  Set cAjustarDesk.ParentForm = Me
  cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
  iX = Read_INI("Configuration", "ExitButtonX", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
  iY = Read_INI("Configuration", "ExitButtonY", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
  cWindows.ButtonExitXY CLng(iX), CLng(iY)
  cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
  cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\config.ini")
  
  cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\VIS_STUDIO\"
  picSpectrum.Left = cWindows.AreaLeft
  picSpectrum.Top = cWindows.AreaTop
  picSpectrum.Width = cWindows.AreaWidth
  picSpectrum.Height = cWindows.AreaHeight
  Label.Width = cWindows.AreaWidth

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  bolVisShow = False
  Cancel = 1
End Sub


Sub Siguiente_Visualizacion()
  On Error Resume Next
  If frmOpciones.cboVisualizacion.ListCount <= 1 Then Exit Sub
  
  IndexVisualization = IndexVisualization + 1
  
  If IndexVisualization >= frmOpciones.cboVisualizacion.ListCount Then IndexVisualization = 0
  
  Load_Visualizacion frmOpciones.cboVisualizacion.List(IndexVisualization)
  Label.Visible = True
  Label.Caption = frmOpciones.cboVisualizacion.List(IndexVisualization)
  Timer.Interval = 2000
  Timer.Enabled = True
End Sub

Sub Anterior_Visualizacion()
  On Error Resume Next
  If frmOpciones.cboVisualizacion.ListCount <= 1 Then Exit Sub
  
  IndexVisualization = IndexVisualization - 1
  
  If IndexVisualization < 0 Then IndexVisualization = frmOpciones.cboVisualizacion.ListCount - 1
  
  Load_Visualizacion frmOpciones.cboVisualizacion.List(IndexVisualization)
  Label.Visible = True
  Label.Caption = frmOpciones.cboVisualizacion.List(IndexVisualization)
  Timer.Interval = 2000
  Timer.Enabled = True
End Sub

Private Sub Timer_Timer()
  Label.Visible = False
  Timer.Enabled = False
End Sub


