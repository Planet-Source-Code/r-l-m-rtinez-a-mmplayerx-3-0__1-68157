VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPopUp 
   Caption         =   "MMPlayerX v. 2.0"
   ClientHeight    =   0
   ClientLeft      =   60
   ClientTop       =   135
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Popup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   0
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   2610
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.FileListBox FileSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00400000&
      Height          =   420
      Hidden          =   -1  'True
      Left            =   1140
      MousePointer    =   99  'Custom
      Pattern         =   "*.mp3;*.wma"
      System          =   -1  'True
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.DirListBox DirSearch 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   165
      TabIndex        =   3
      Top             =   225
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.ListBox lstLanguage 
      Height          =   645
      ItemData        =   "Popup.frx":000C
      Left            =   105
      List            =   "Popup.frx":000E
      TabIndex        =   2
      Top             =   3300
      Width           =   7470
   End
   Begin VB.PictureBox picDefaultLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   1110
      Picture         =   "Popup.frx":0010
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   338
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   5070
   End
   Begin VB.FileListBox fileBmps 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   225
      Hidden          =   -1  'True
      Left            =   1935
      Pattern         =   "*.jpg;*.bmp"
      System          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuMenuPrincipal 
      Caption         =   "MenuPrincipal"
      Begin VB.Menu mnuPlay 
         Caption         =   "Reproducir"
         Begin VB.Menu mnuArchivo 
            Caption         =   "Archivo..."
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuFolder 
            Caption         =   "Folder..."
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuNuevaBusqueda 
            Caption         =   "Buscar Audio ..."
         End
      End
      Begin VB.Menu mnuA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVentanas 
         Caption         =   "Ventanas"
         Begin VB.Menu mnuListaR 
            Caption         =   "Lista Reproduccion"
         End
         Begin VB.Menu mnuBibliotecaMultimedia 
            Caption         =   "Biblioteca Mutimedia"
         End
         Begin VB.Menu mnuEqualizador 
            Caption         =   "Equalizador"
         End
         Begin VB.Menu mnuMaxSpec 
            Caption         =   "Show Visualization"
         End
         Begin VB.Menu mnuMCaratula 
            Caption         =   "Maximizar Caratula"
         End
         Begin VB.Menu mnuTagEditor 
            Caption         =   "Tag Editor"
         End
         Begin VB.Menu mnuLyrics 
            Caption         =   " Lyrics"
         End
      End
      Begin VB.Menu mnuxy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControles 
         Caption         =   "Controles de Reproduccion"
         Begin VB.Menu mnuVolumen 
            Caption         =   "   Volumen"
            Begin VB.Menu mnuSubirVolumen 
               Caption         =   "+   Subir Volumen"
            End
            Begin VB.Menu mnuBajarVolumen 
               Caption         =   "-   Bajar Volumen"
            End
         End
         Begin VB.Menu mnuD 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTrackAnterior 
            Caption         =   "Z   Track Anterior"
         End
         Begin VB.Menu mnuReproducir 
            Caption         =   "X   Reproducir"
         End
         Begin VB.Menu mnuPausa 
            Caption         =   "C   Pausa"
         End
         Begin VB.Menu mnuDetener 
            Caption         =   "V   Detener"
         End
         Begin VB.Menu mnuSigTrack 
            Caption         =   "B   Siguiente Track"
         End
         Begin VB.Menu mnuF 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIntro 
            Caption         =   "I   Intro 10 Segundos"
         End
         Begin VB.Menu mnuRepetir 
            Caption         =   "R   Repetir Track"
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "S   Silencio"
         End
         Begin VB.Menu mnuOrdenAleatorio 
            Caption         =   "W   Orden Aleatorio"
         End
         Begin VB.Menu mnuG 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAtras5Seg 
            Caption         =   "A   Atras 5 Segundos"
         End
         Begin VB.Menu mnuAdelante5Seg 
            Caption         =   "D   Adelante 5 Segundos"
         End
      End
      Begin VB.Menu mnuh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones 
         Caption         =   "Opciones ..."
      End
      Begin VB.Menu mnuSkins 
         Caption         =   "Skins"
         WindowList      =   -1  'True
         Begin VB.Menu mnuExpSkins 
            Caption         =   "<<  Explorador de Skins >>"
         End
         Begin VB.Menu mnuxxx 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSkinsAdd 
            Caption         =   ""
            Index           =   1
         End
      End
      Begin VB.Menu mnuWOpacity 
         Caption         =   "Window Opacity"
         Begin VB.Menu mnuAlpha 
            Caption         =   "100%"
            Index           =   0
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "90%"
            Index           =   1
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "80%"
            Index           =   2
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "70%"
            Index           =   3
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "60%"
            Index           =   4
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "50%"
            Index           =   5
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "40%"
            Index           =   6
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "30%"
            Index           =   7
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "20%"
            Index           =   8
         End
         Begin VB.Menu mnuAlpha 
            Caption         =   "10%"
            Index           =   9
         End
         Begin VB.Menu mnuAlphaPer 
            Caption         =   "Personalizar..."
         End
      End
      Begin VB.Menu mnui 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "Acerca de ..."
      End
      Begin VB.Menu mnuxx 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuMainSpec 
      Caption         =   "MainSpectrum"
      Begin VB.Menu mnuSpecNone 
         Caption         =   "None Visualisation"
      End
      Begin VB.Menu mnuSpecBars 
         Caption         =   "Spectrum Analyzer"
      End
      Begin VB.Menu mnuSpecOsc 
         Caption         =   "Oscilloscope"
      End
   End
   Begin VB.Menu mnuSpectrum 
      Caption         =   "Spectrum"
      Begin VB.Menu mnuPrevVis 
         Caption         =   "Previous Visualization"
      End
      Begin VB.Menu mnuNextVis 
         Caption         =   "Next Visualization"
      End
      Begin VB.Menu mnuConfigVis 
         Caption         =   "Configure Visualization"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMainLista 
      Caption         =   "Lista Reproduccion"
      Begin VB.Menu mnuQuitar 
         Caption         =   "Quitar"
         Begin VB.Menu mnuQuitarS 
            Caption         =   "Quitar Seleccionados"
         End
         Begin VB.Menu mnuLimpiarL 
            Caption         =   "Limpiar Lista"
         End
         Begin VB.Menu mnuQuitarAB 
            Caption         =   "Quitar archivo Borrados"
         End
      End
      Begin VB.Menu mnuMis 
         Caption         =   "Mis"
         Begin VB.Menu mnuOrdenarTitulo 
            Caption         =   "Order por Titulo"
         End
         Begin VB.Menu mnuOrdenarArtista 
            Caption         =   "Ordenar por Artista"
         End
         Begin VB.Menu mnuOrdenarArchivo 
            Caption         =   "Ordenar por Archivo"
         End
         Begin VB.Menu mnuEditarArchivos 
            Caption         =   "Editar Archivo(s)"
         End
      End
      Begin VB.Menu mnuLista 
         Caption         =   "Lista"
         Begin VB.Menu mnuAbrirL 
            Caption         =   "Abrir lista de Biblioteca"
            Begin VB.Menu mnuListaAdd 
               Caption         =   ""
               Index           =   0
            End
         End
         Begin VB.Menu mnuGuardarL 
            Caption         =   "Guardar Lista Reproduccion"
         End
         Begin VB.Menu mnuAdministrador 
            Caption         =   "Administrador de Listas"
         End
      End
   End
   Begin VB.Menu mnuBibliotecaM 
      Caption         =   "Menu Biblioteca"
      Begin VB.Menu mnuBiblioteca 
         Caption         =   "biblioteca"
         Begin VB.Menu mnuRemoverbiblioteca 
            Caption         =   "Remover Archivos eliminados de la biblioteca"
         End
         Begin VB.Menu mnuAgregarAudio 
            Caption         =   "Agregar Audio a la Biblioteca"
         End
      End
      Begin VB.Menu mnuMenubiblioteca 
         Caption         =   "MenuBiblioteca"
         Begin VB.Menu mnuReproducirSeleccionadosAgregar 
            Caption         =   "Reproducir Seleccionados Agregar"
         End
         Begin VB.Menu mnuReproducirTodosAgregar 
            Caption         =   "Reproducir Todos Agregar"
         End
         Begin VB.Menu mnu1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuReproducirSeleccionadosNuevo 
            Caption         =   "Reproducir Seleccionados Nuevo"
         End
         Begin VB.Menu mnuReproducirTodosNuevo 
            Caption         =   "Reproducir Todos Nuevo"
         End
         Begin VB.Menu mnu2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGuardarTodoLista 
            Caption         =   "Guardar Todo como lista"
         End
         Begin VB.Menu mnuEditarInformacionArchivos 
            Caption         =   "Editar Informacion de Archivos"
         End
         Begin VB.Menu mnu3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExplorarCarpeta 
            Caption         =   "Explorar Carpeta"
         End
         Begin VB.Menu mnu4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRemoverbibliotecaM 
            Caption         =   "Remover de la Biblioteca"
         End
      End
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_GotFocus()
    ' When the user presses the taskbar button, this form gets
    ' the focus, so we shift the focus to the frmPad
    'frmMain.SetFocus
End Sub

Private Sub Form_Load()
    ' The wrapper window is practically invisible to the user.
    ' Older version of WindowBlinds have a bug that causes this
    ' window to appear.
    Me.Top = -10000
    Me.Icon = frmMain.Icon
    Cargar_Lista_Rep
End Sub

Sub Cargar_Lista_Rep()
     '-----------------------------------------------------------------------------------
    '// buskar los archivos de playlist y agragarlos
Dim i As Integer

If Dir(tAppConfig.AppConfig & "Library\", vbDirectory) <> "" Then
  fileBmps.Path = tAppConfig.AppConfig & "Library\"
  fileBmps.Pattern = "*.pls"
  
  frmPopUp.mnuListaAdd(0).Caption = ""
  For i = 1 To frmPopUp.mnuListaAdd.count - 1
    Unload frmPopUp.mnuListaAdd(i)
  Next i
  
  For i = 0 To fileBmps.ListCount - 1
    If i > frmPopUp.mnuListaAdd.count - 1 Then Load frmPopUp.mnuListaAdd(i)
           
    frmPopUp.mnuListaAdd(i).Caption = " " & Left(fileBmps.List(i), Len(fileBmps.List(i)) - 4)
    frmPopUp.mnuListaAdd(i).Visible = True
  Next i
End If

End Sub

Private Sub Form_Resize()
    ' Change frmmain's state according to changes made to this
    ' form using the taskbar
    
 If bLoading = True Then Exit Sub
   
    If Me.WindowState = vbMinimized Then
           If bolCaratulaShow = True Then frmCaratula.Hide
           If bolOpcionesShow = True Then frmOpciones.Hide
           If bolAcercaShow = True Then frmAcerca.Hide
           If bolTagsShow = True Then frmTags.Hide
           If bolLyricsShow = True Then frmLyrics.Hide
           If bolSplashScreen = True Then frmSplash.Hide
           If bolVisShow = True Then frmSpectrum.Hide
           If bolSearchShow = True Then frmSearch.Hide
           If bolPlayListShow = True Then frmPlayList.Hide
           If bolMediaLibraryShow = True Then frmLibrary.Hide
           frmMain.Hide
    Else
           If bolAcercaShow = True Then frmAcerca.Show
           If bolCaratulaShow = True Then frmCaratula.Show
           If bolOpcionesShow = True Then frmOpciones.Show
           If bolLyricsShow = True Then frmLyrics.Show
           If bolTagsShow = True Then frmTags.Show
           If bolVisShow = True Then frmSpectrum.Show
           If bolSearchShow = True Then frmSearch.Show
           If bolPlayListShow = True Then frmPlayList.Show
           If bolMediaLibraryShow = True Then frmLibrary.Show
           
        frmMain.WindowState = Me.WindowState
        frmMain.Visible = True
        If bolSplashScreen = True Then frmSplash.Show
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload frmMain
 End
End Sub

Private Sub mnuAcercaDe_Click()
 If bolAcercaShow = True Then
   frmAcerca.ZOrder 0
 Else
   frmAcerca.Show
 End If
End Sub


Private Sub mnuAdelante5Seg_Click()
 frmMain.Five_Seg_Forward
End Sub

Private Sub mnuAdministrador_Click()
  frmLibrary.TreeFiles.Nodes("kPlaL").Selected = True
  frmLibrary.Show

End Sub

Private Sub mnuAgregarAudio_Click()
frmSearch.Show
End Sub

Private Sub mnuAlpha_Click(Index As Integer)
On Error GoTo HELL
 Dim tAlpha
 Dim i As Integer
   tAlpha = mnuAlpha(Index).Caption
   tAlpha = Left(tAlpha, Len(tAlpha) - 1)
  Call SetWindowLong(frmMain.hwnd, GWL_EXSTYLE, GetWindowLong(frmMain.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
  Call SetLayeredWindowAttributes(frmMain.hwnd, 0, (255 * tAlpha) / 100, LWA_ALPHA)
  mnuAlpha(Index).Checked = True
  Call SetWindowLong(frmPlayList.hwnd, GWL_EXSTYLE, GetWindowLong(frmMain.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
'  Call SetLayeredWindowAttributes(frmPlayList.hwnd, 0, (255 * tAlpha) / 100, LWA_ALPHA)

  OpcionesMusic.Alpha = tAlpha
    
    frmPopUp.mnuAlphaPer.Caption = LineLanguage(37)
    frmPopUp.mnuAlphaPer.Checked = False
  For i = 0 To 9
   If i <> Index Then mnuAlpha(i).Checked = False
  Next i
 Exit Sub
HELL:
End Sub

Private Sub mnuAlphaPer_Click()
   frmOpciones.Show
   frmOpciones.Select_Option 1
   frmOpciones.TSAppConfig.Tabs(3).Selected = True
End Sub


Private Sub mnuArchivo_Click()
 Dialogo.Filter = "Audio Files |" & strPathern
 Dialogo.ShowOpen

 If Dir(Dialogo.filename) = "" Or Dialogo.filename = "" Then Exit Sub
 frmPlayList.Agregar_PlayList_de_Archivo Dialogo.filename


End Sub

Private Sub mnuAtras5Seg_Click()
 frmMain.Five_Seg_Backward
End Sub

Private Sub mnuBajarVolumen_Click()
 '// bajar volumen
  frmMain.Form_KeyPress 45
End Sub

Private Sub mnuBibliotecaM_Click()
frmMain.Mostrar_Media_Library
End Sub

Private Sub mnuBibliotecaMultimedia_Click()
  frmMain.Mostrar_Media_Library
End Sub

Private Sub mnuConfigVis_Click()
 On Error Resume Next
   frmOpciones.cboVisualizacion.ListIndex = IndexVisualization
   frmOpciones.Select_Option 8
   frmOpciones.Show

End Sub

Private Sub mnuDetener_Click()
 frmMain.Stop_Player
End Sub

Private Sub mnuEditarArchivos_Click()
frmPlayList.Editar_Archivos
End Sub

Private Sub mnuEditarInformacionArchivos_Click()
frmLibrary.Editar_Archivos
End Sub

Private Sub mnuEqualizador_Click()
   frmOpciones.Select_Option 7
   frmOpciones.Show
End Sub

Private Sub mnuExit_Click()
 frmSpectrum.Hide
 bolVisShow = False
End Sub


Private Sub mnuExplorarCarpeta_Click()
frmLibrary.Explorar_Archivos
End Sub

Private Sub mnuExpSkins_Click()
   frmOpciones.Select_Option 2
   frmOpciones.Show
End Sub

Private Sub mnuFolder_Click()
 On Error GoTo HELL
 Dim sPath As String
  sPath = Explorador_Para_Directorios(Me.hwnd, LineLanguage(234))

  If sPath = "" Then Exit Sub
  Search_Files sPath

Exit Sub
HELL:

End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  BUSKEDA METODO UNO: MAS RAPIDO PERO UTILIZANDO OBJETOS DIR Y FILE :)                 |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Search_Files(strPath As String)
 On Error GoTo HELL
 Dim strPathCur As String
 Dim bEncontro As Boolean
 Dim i As Integer
 '// Primero buscar en el directorio padre para buscar despues en subdirectorios
 If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
 
 
 '// set pather at files list box
 If strPathern = "" Then strPathern = "*.mp3"
   
  FileSearch.Pattern = strPathern
  
  FileSearch.Path = strPath
 If FileSearch.ListCount > 0 Then
    For i = 0 To FileSearch.ListCount - 1
        frmPlayList.Agregar_PlayList_de_Archivo FileSearch.Path & "\" & FileSearch.List(i)
    Next
 End If
 '// poner cursor de busqueda si hay del skin
 strPathCur = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\"
 If Dir(strPathCur & "curFind.cur") <> "" Then
   frmMain.picNormalMode.MouseIcon = LoadPicture(strPathCur & "curFind.cur")
 End If
  
 '// Empezar ha buskar
 bSearching = True
 Call Start_Search(strPath)
 bSearching = False
  
HELL:
 If Dir(strPathCur & "curMain.cur") <> "" Then frmMain.picNormalMode.MouseIcon = LoadPicture(strPathCur & "curMain.cur")
End Sub


'// metod for search is very faster

Sub Start_Search(strPath As String)
 On Error Resume Next  '// manejador de error por si permisos de acceso a los directorios
 
 DoEvents '// para que deje trabajar el Windows
 Dim subdirs As Integer, k As Integer, intFolder As Integer
 ReDim subdirs_name(0 To 10) As String  '// arreglo para directorios
 Dim i As Integer
 subdirs = 0

If bSearching = False Then Exit Sub  '// para cancelar si keremos
  
 '// Poner el Dir en la direccion para iniciar busqueda y en subdirectorios
 DirSearch.Path = strPath
For intFolder = 0 To DirSearch.ListCount - 1  '// buskar en los elementos del dir
      '// Komo todos son directorios almacenarlos en el arreglo para despues buskar
      '// en subdirectorios
      subdirs_name(subdirs) = DirSearch.List(intFolder)
      subdirs = subdirs + 1
      '// si se pasan los directorios del maximo del arreglo
      '// aumentar otros 10
      If subdirs Mod 10 = 0 Then ReDim Preserve subdirs_name(0 To subdirs + 10)
      
      '// Verifikar si hay mp3s con el file
      FileSearch.Path = DirSearch.List(intFolder)
      If FileSearch.ListCount > 0 Then
        '// Ir kontando todos los mp3's
        For i = 0 To FileSearch.ListCount - 1
            frmPlayList.Agregar_PlayList_de_Archivo FileSearch.Path & "\" & FileSearch.List(i)
        Next

      End If
Next intFolder

'//-----------Buscamos en subdirectorios ----------------------------------------
'// como es una procedimento que se llama a si mismo las variables anteriores
'// se siguen conservando hasta que termine
For k = 0 To subdirs - 1
 Start_Search subdirs_name(k)
Next

End Sub


Private Sub mnuGuardarL_Click()
frmPlayList.Guardar_PlayList
End Sub

Private Sub mnuGuardarTodoLista_Click()
 frmLibrary.Guardar_PlayList True
End Sub

Private Sub mnuIntro_Click()
  frmPopUp.mnuIntro.Checked = Not frmPopUp.mnuIntro.Checked
  frmMain.Intro
End Sub

Private Sub mnuLimpiarL_Click()
 frmPlayList.PL.Clear
 frmPlayList.PL.Rows = 0
End Sub

Private Sub mnuListaAdd_Click(Index As Integer)

 Dim sTrack As String
 Dim sFilePL As String
 Dim i As Integer
 
 
 sFilePL = tAppConfig.AppConfig & "Library\" & Trim(mnuListaAdd(Index).Caption) & ".pls"

 If Dir(sFilePL) <> "" Then
    frmPlayList.PL.Clear
    frmPlayList.PL.Rows = 0
    frmLibrary.TreeFiles.Nodes("kPlaE" & Trim(mnuListaAdd(Index).Caption)).Selected = True
    frmLibrary.TreeFiles_Click
    frmLibrary.Agregar_Todos True
 End If

End Sub

Private Sub mnuListaR_Click()
frmMain.Mostrar_Play_List
End Sub

Private Sub mnuLyrics_Click()
 bolLyricsShow = Not bolLyricsShow
 frmLyrics.Visible = bolLyricsShow
 frmPopUp.mnuLyrics.Checked = bolLyricsShow
End Sub

Private Sub mnuMaxSpec_Click()
    frmMain.Mostrar_Visualizacion
End Sub

Private Sub mnuMCaratula_Click()
 bolCaratulaShow = Not bolCaratulaShow
 frmCaratula.Visible = bolCaratulaShow
 frmPopUp.mnuMCaratula.Checked = bolCaratulaShow
End Sub


Private Sub mnuNextVis_Click()
frmSpectrum.Siguiente_Visualizacion
End Sub

Private Sub mnuNuevaBusqueda_Click()
  frmSearch.Show
End Sub

Private Sub mnuOpciones_Click()
   frmOpciones.Select_Option 1
   frmOpciones.Show
 
End Sub

Private Sub mnuOrdenAleatorio_Click()
frmMain.Randomize_Click
End Sub

Private Sub mnuOrdenarArchivo_Click()
frmPlayList.PL.Col = 10
frmPlayList.PL.Sort = flexSortStringAscending
End Sub

Private Sub mnuOrdenarArtista_Click()
frmPlayList.PL.Col = 2
frmPlayList.PL.Sort = flexSortStringAscending

End Sub

Private Sub mnuOrdenarTitulo_Click()
frmPlayList.PL.Col = 0
frmPlayList.PL.Sort = flexSortStringAscending

End Sub

Private Sub mnuPausa_Click()
 frmMain.Pause_Play
End Sub

Private Sub mnuPrevVis_Click()
frmSpectrum.Anterior_Visualizacion
End Sub

Private Sub mnuQuitarAB_Click()
frmPlayList.Remover_Archivos_Eliminados
End Sub

Private Sub mnuQuitarS_Click()
frmPlayList.Remover_Tracks
End Sub

Private Sub mnuRemoverbiblioteca_Click()
frmLibrary.Eliminar_Archivos_Biblioteca
End Sub

Private Sub mnuRemoverbibliotecaM_Click()
frmLibrary.Eliminar_Biblioteca
End Sub

Private Sub mnuRepetir_Click()
 frmPopUp.mnuRepetir.Checked = Not frmPopUp.mnuRepetir.Checked
 frmMain.Player_Repeat
End Sub

Private Sub mnuReproducir_Click()
 frmMain.Play
End Sub

Private Sub mnuReproducirSeleccionadosAgregar_Click()
frmLibrary.Agregar_Seleccionadas False
End Sub

Private Sub mnuReproducirSeleccionadosNuevo_Click()
frmLibrary.Agregar_Seleccionadas True
End Sub

Private Sub mnuReproducirTodosAgregar_Click()
frmLibrary.Agregar_Todos False
End Sub

Private Sub mnuReproducirTodosNuevo_Click()
frmLibrary.Agregar_Todos True
End Sub

Private Sub mnuSalir_Click()
 Unload frmMain
End Sub

Private Sub mnuSigTrack_Click()
 frmMain.Next_Track
End Sub

Private Sub mnuSilencio_Click()
  frmPopUp.mnuSilencio.Checked = Not frmPopUp.mnuSilencio.Checked
  frmMain.Player_Mute
End Sub

Private Sub mnuSkinsAdd_Click(Index As Integer)
 On Error Resume Next
 Dim Skins As String, MiRuta As String
 Dim i As Integer
 Skins = Trim(mnuSkinsAdd(Index).Caption)
 '// si es el mismo skin salir
 If Skins = "" Then Exit Sub
 
 '// chekar si existe la carpeta
 MiRuta = tAppConfig.AppConfig & "Skins\"
 If Dir(MiRuta & Skins, vbDirectory) = "" Then Exit Sub

 If LCase(Skins) = LCase(tAppConfig.Skin) Then Exit Sub
   '// seleccionar el skin
   For i = 1 To mnuSkinsAdd.count
      If i = Index Then
        mnuSkinsAdd(Index).Checked = True
      Else
        mnuSkinsAdd(i).Checked = False
      End If
   Next i
 

    frmMain.Visible = False

    '// Cambiar el skin
    Change_Skin Skins
    
    '// CHANGE WINDOWS
 frmPlayList.cargar_formulario
 frmLibrary.cargar_formulario
 frmSpectrum.cargar_formulario
 frmCaratula.cargar_formulario
 frmLyrics.cargar_formulario

    
    '// ajustar los bordes
    Form_Mini_Normal
    
    Change_Mask bMiniMask, False
    
    frmOpciones.lblSkin(2).Caption = Skins
    frmOpciones.ListaSkins.Selected(Index) = True
    frmOpciones.ListaSkins.ListIndex = Index - 1
    
    frmPlayList.Show_ScrollBar

    frmMain.Visible = True

End Sub

Public Sub mnuSpecBars_Click()
With frmPopUp
   .mnuSpecBars.Checked = True
   .mnuSpecNone.Checked = False
   .mnuSpecOsc.Checked = False
 End With
End Sub

Public Sub mnuSpecNone_Click()
 With frmPopUp
   .mnuSpecBars.Checked = False
   .mnuSpecNone.Checked = True
   .mnuSpecOsc.Checked = False
 End With
End Sub

Public Sub mnuSpecOsc_Click()
With frmPopUp
   .mnuSpecBars.Checked = False
   .mnuSpecNone.Checked = False
   .mnuSpecOsc.Checked = True
 End With
End Sub

Private Sub mnuSubirVolumen_Click()
 frmMain.Form_KeyPress 43 '// subir volumen

End Sub

Private Sub mnuTagEditor_Click()
 If bolTagsShow = True Then
'    frmTags.Load_Tags
    frmTags.ZOrder 0
 Else
   frmTags.Show
 End If
End Sub

Private Sub mnuTrackAnterior_Click()
 frmMain.Previous_Track
End Sub

