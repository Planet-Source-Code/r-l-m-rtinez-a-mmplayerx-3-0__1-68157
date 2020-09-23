Attribute VB_Name = "mLanguage"
Option Explicit

'Public arryLanguage() As String

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  IDIOMA                                                                               |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Load_Language_Spanish()
With frmPopUp
.lstLanguage.Clear
.lstLanguage.AddItem "Language"
' MENU
.lstLanguage.AddItem " Reproducir"  ' 1
.lstLanguage.AddItem "  Archivo..." ' 2
.lstLanguage.AddItem "  Folder..." ' 3
.lstLanguage.AddItem "  Nueva Busqueda" ' 4
.lstLanguage.AddItem " Ventanas" '  5
.lstLanguage.AddItem "  Lista de Reproduccion" '  6
.lstLanguage.AddItem "  Biblioteca Multimedia" ' 7
.lstLanguage.AddItem "  Equalizador" ' 8
.lstLanguage.AddItem "  Mostrar Visualizacion" '  9
.lstLanguage.AddItem "  Maximizar Caratula" ' 10
.lstLanguage.AddItem "  Editar Track(s) Tag" ' 11
.lstLanguage.AddItem "  Karaoke" ' 12
.lstLanguage.AddItem " Controles de Reproducción" ' 13
.lstLanguage.AddItem "   Volumen" ' 14
.lstLanguage.AddItem "+     Subir Volumen" ' 15
.lstLanguage.AddItem "-     Bajar Volumen" ' 16
.lstLanguage.AddItem "Z   Track Anterior" ' 17
.lstLanguage.AddItem "X   Reproducir" ' 18
.lstLanguage.AddItem "C   Pausar" ' 19
.lstLanguage.AddItem "V   Detener" ' 20
.lstLanguage.AddItem "B   Siguiente Track" ' 21
.lstLanguage.AddItem "I   Intro 10 seg." ' 22
.lstLanguage.AddItem "R   Repetir Track" ' 23
.lstLanguage.AddItem "S   Silencio" ' 24
.lstLanguage.AddItem "   Orden aleatorio" ' 25
.lstLanguage.AddItem "A   Atras 5 seg." ' 26
.lstLanguage.AddItem "D   Adelante 5 Seg." ' 27
.lstLanguage.AddItem " Opciones" ' 28
.lstLanguage.AddItem " Skins" ' 29
.lstLanguage.AddItem "   << Explorador de Skins >>" ' 30
.lstLanguage.AddItem " Transparencia" ' 31
.lstLanguage.AddItem "   Personalizar" ' 32
.lstLanguage.AddItem " Acerca de" ' 33
.lstLanguage.AddItem " Salir" ' 34
' ACERCA
.lstLanguage.AddItem " Acerca de MMPlayerX" ' 35
' CARATULA
.lstLanguage.AddItem "Caratula actual" ' 36
' KARAOKE
.lstLanguage.AddItem "Karaoke" ' 37
.lstLanguage.AddItem "  [ Letras no Encontradas ]" ' 38
' MAIN
.lstLanguage.AddItem "    Menu" ' 39
.lstLanguage.AddItem "    Minimizar"
.lstLanguage.AddItem "    Change Mode" ' 41
.lstLanguage.AddItem "    Salir" ' 42
.lstLanguage.AddItem "  Sin Visualización" '43
.lstLanguage.AddItem "  Analizador de Espectro" '44
.lstLanguage.AddItem "  Osiloscopio" '45
.lstLanguage.AddItem ""
.lstLanguage.AddItem ""
.lstLanguage.AddItem ""
.lstLanguage.AddItem ""
.lstLanguage.AddItem ""
.lstLanguage.AddItem ""
.lstLanguage.AddItem ""
.lstLanguage.AddItem ""
.lstLanguage.AddItem ""
.lstLanguage.AddItem ""
.lstLanguage.AddItem ""
' TAGS
.lstLanguage.AddItem "Editor de Tags + Información MPEG" ' 57
.lstLanguage.AddItem "  Multiples tracks estan seleccionados, Selecciona los checkboxs para aplicar los cambios  a TODOS los archivos seleccionados."
.lstLanguage.AddItem "  Seleccionar Todo" '58
.lstLanguage.AddItem "  Tags" '59
.lstLanguage.AddItem "  Karaoke"
.lstLanguage.AddItem "  Agregar" '61
.lstLanguage.AddItem "  Deshacer"
.lstLanguage.AddItem "  Aceptar"
.lstLanguage.AddItem "  Cancelar" '64
.lstLanguage.AddItem "  Aplicar"
.lstLanguage.AddItem "  Caratula" '66
.lstLanguage.AddItem "   Agregar C."
.lstLanguage.AddItem "   Remover C." '68

' OPCIONES
.lstLanguage.AddItem "Opciones" ' 70
.lstLanguage.AddItem "  Aceptar" ' 71
.lstLanguage.AddItem "  Cancelar" ' 72
.lstLanguage.AddItem "  Aplicar" ' 73
.lstLanguage.AddItem "Aplicación" ' 74
.lstLanguage.AddItem "  Trayectoria Configuración" ' 75
.lstLanguage.AddItem "    Trayectoria de skins y configuración:" ' 76
.lstLanguage.AddItem "    Explorar..." ' 77
.lstLanguage.AddItem "    Nota: Algunas opciones requieren que se reinicie la aplicación."
.lstLanguage.AddItem "    Memoria Libre (Fisica):" ' 79
.lstLanguage.AddItem "  Aplicación" ' 80
.lstLanguage.AddItem "    Lenguaje:" ' 81
.lstLanguage.AddItem "    Siempre arriba." ' 82
.lstLanguage.AddItem "    Mostrar Splash Screen." ' 83
.lstLanguage.AddItem "    Permitir multiples instancias." ' 84
.lstLanguage.AddItem "    Habilitar menu en drives y directorios." ' 85
.lstLanguage.AddItem "    Mostrar MMPlayerX en:" ' 86
.lstLanguage.AddItem "    Barra de tareas." ' 87
.lstLanguage.AddItem "    Bandeja de sistema." ' 88
.lstLanguage.AddItem "  Transparencia" ' 89
.lstLanguage.AddItem "    Transparencia(Solo win 2000 o sup.)" ' 90
.lstLanguage.AddItem "Skins" ' 91
.lstLanguage.AddItem "  Skin actual:" ' 92
.lstLanguage.AddItem "  Información:" ' 93
.lstLanguage.AddItem "  Cargar región desde archivo." ' 94
.lstLanguage.AddItem "Wallpaper" ' 95
.lstLanguage.AddItem "  Opciones de fondo de escritorio." ' 96
.lstLanguage.AddItem "  No alterar." ' 97
.lstLanguage.AddItem "  Ajustar." ' 98
.lstLanguage.AddItem "  Centrar." ' 99
.lstLanguage.AddItem "  Mosaico." ' 100
.lstLanguage.AddItem "  Proporcional." ' 101
.lstLanguage.AddItem "Play List" ' 102
.lstLanguage.AddItem "  Formato Lista de Reproducción." ' 103
.lstLanguage.AddItem "  Formato de texto reproduciendo." ' 104
.lstLanguage.AddItem "  Tipo de Scroll:" ' 105
.lstLanguage.AddItem "    Rotar." ' 106
.lstLanguage.AddItem "    Zig Zag." ' 107
.lstLanguage.AddItem "  Velocidad del Scroll:" ' 108
.lstLanguage.AddItem "Reproductor" ' 109
.lstLanguage.AddItem "  Reproducir archivos:" ' 110
.lstLanguage.AddItem "  Mostrar icono en bandeja de sistema:" ' 111
.lstLanguage.AddItem "    Anterior Track." ' 112
.lstLanguage.AddItem "    Reproducir." ' 113
.lstLanguage.AddItem "    Pausar." ' 114
.lstLanguage.AddItem "    Detener." ' 115
.lstLanguage.AddItem "    Siguente Track." ' 116
.lstLanguage.AddItem "  Crossfade entre tracks (ms):" ' 117
.lstLanguage.AddItem "  Crossfade en Detener (ms):" ' 118
.lstLanguage.AddItem "  Reproducir al inicio." ' 119
.lstLanguage.AddItem "Efectos FX" ' 120
.lstLanguage.AddItem "  coro" ' 121
.lstLanguage.AddItem "    Habilitar coro." ' 122
.lstLanguage.AddItem "      Mezcla:" ' 123
.lstLanguage.AddItem "      Profundidad:" ' 124
.lstLanguage.AddItem "      Retroacción:" ' 125
.lstLanguage.AddItem "      Frecuencía:" ' 126
.lstLanguage.AddItem "      Forma de onda:" ' 127
.lstLanguage.AddItem "      Retrazo:" ' 128
.lstLanguage.AddItem "      Fase:" ' 129
.lstLanguage.AddItem "  compresor" ' 130
.lstLanguage.AddItem "    Habilitar compresor." ' 131
.lstLanguage.AddItem "      Incremento:" ' 132
.lstLanguage.AddItem "      Ataque:"
.lstLanguage.AddItem "      Edición:"
.lstLanguage.AddItem "      Umbral:"
.lstLanguage.AddItem "      Proporción:"
.lstLanguage.AddItem "      Preretrazo:"
.lstLanguage.AddItem "  Distorción"
.lstLanguage.AddItem "    Habilitar Distorción:"
.lstLanguage.AddItem "      Incremento:" ' 140
.lstLanguage.AddItem "      Bordes:"
.lstLanguage.AddItem "      frecuencia Central:"
.lstLanguage.AddItem "      Ancho frecuencia:"
.lstLanguage.AddItem "      Atenuación:"
.lstLanguage.AddItem "  eco" '145
.lstLanguage.AddItem "    Habilitar eco."
.lstLanguage.AddItem "      Mezcla:"
.lstLanguage.AddItem "      Retroaccion:"
.lstLanguage.AddItem "      Atraso izquierda:"
.lstLanguage.AddItem "      Atraso derecha:"
.lstLanguage.AddItem "      Atraso Central:"
.lstLanguage.AddItem "  Flanger" ' 152
.lstLanguage.AddItem "    Habilitar Flanger."
.lstLanguage.AddItem "      Mezcla:"
.lstLanguage.AddItem "      Profundidad:"
.lstLanguage.AddItem "      Retroaccion:"
.lstLanguage.AddItem "      Frecuencia:"
.lstLanguage.AddItem "      Forma de Onda:"
.lstLanguage.AddItem "      Retrazo:"
.lstLanguage.AddItem "      Fase:" '160
.lstLanguage.AddItem "  Gargarizar" ' 161
.lstLanguage.AddItem "    Habilitar gargarizar."
.lstLanguage.AddItem "      Hz:"
.lstLanguage.AddItem "      Forma de Onda:"
.lstLanguage.AddItem "  I3DL2 Reverberación" ' 165
.lstLanguage.AddItem "    Habilitar I3D nivel 2 Reverberación."
.lstLanguage.AddItem "      Cuarto:"
.lstLanguage.AddItem "      Cuarto HF:"
.lstLanguage.AddItem "      Factor giratorio:" '169
.lstLanguage.AddItem "      Tiempo decadencia:"
.lstLanguage.AddItem "      Prop. dec. HF:"
.lstLanguage.AddItem "      Reflecciones:"
.lstLanguage.AddItem "      Atraso Refleccción:"
.lstLanguage.AddItem "      Reverberación:"
.lstLanguage.AddItem "      Atraso de Rev.:"
.lstLanguage.AddItem "      Difusión:"
.lstLanguage.AddItem "      Densidad:"
.lstLanguage.AddItem "      HF Referencia:"
.lstLanguage.AddItem "  Reverberación" '179
.lstLanguage.AddItem "    Habilitar Reverberación de ondas."
.lstLanguage.AddItem "      Incremento:"
.lstLanguage.AddItem "      Mezcla Reverberación:"
.lstLanguage.AddItem "      Tiempo de Rev.:"
.lstLanguage.AddItem "      HF Proporción:"
.lstLanguage.AddItem "  Valores por default" ' 185
.lstLanguage.AddItem "  Desabilitar todos"
.lstLanguage.AddItem "Equalizador"
.lstLanguage.AddItem "  Habilitar EQ."
.lstLanguage.AddItem "  Presentes:" '190
.lstLanguage.AddItem "  Borrar EQ"
.lstLanguage.AddItem "  Guardar EQ"
.lstLanguage.AddItem "  Nombre del Equalizador:"
.lstLanguage.AddItem "  Borrar equalizador:"
.lstLanguage.AddItem "Visualización" '195
.lstLanguage.AddItem "  Visualizaciones:"
.lstLanguage.AddItem "  Presentes:"
.lstLanguage.AddItem "  Nuevos:"
.lstLanguage.AddItem "  Tipo Fondo:" '198
.lstLanguage.AddItem "  Peaks:" '200
.lstLanguage.AddItem "  Barras:"
.lstLanguage.AddItem "  Archivo Imagen:"
.lstLanguage.AddItem "  Escala:"
.lstLanguage.AddItem "  Color Barras:"
.lstLanguage.AddItem "  Num. Barras:" '205
.lstLanguage.AddItem "  Espacio:"
.lstLanguage.AddItem "  Reflejo:"
.lstLanguage.AddItem "  Color Peak:"
.lstLanguage.AddItem "  Alto Peak:"
.lstLanguage.AddItem "  Gravedad Peak:" '210
.lstLanguage.AddItem "  Gradiente:"
.lstLanguage.AddItem "  Color Fondo:"
.lstLanguage.AddItem "  Color Linea:"
.lstLanguage.AddItem "  Num. Lineas"
.lstLanguage.AddItem "  Alineacion:" '215
.lstLanguage.AddItem "  Guardar"
.lstLanguage.AddItem "  Guardar como"
.lstLanguage.AddItem "  Borrar"
.lstLanguage.AddItem "  Mostrar" '219
.lstLanguage.AddItem "  Borrar Visualización:"
.lstLanguage.AddItem "  Nombre de Visualización:"
.lstLanguage.AddItem "  Anterior Visualizacion"
.lstLanguage.AddItem "  Siguiente Visualizacion"
.lstLanguage.AddItem "  Configurar ..." '224
.lstLanguage.AddItem "  Salir"
.lstLanguage.AddItem "  Guardar Config."
.lstLanguage.AddItem "  Física:"
.lstLanguage.AddItem "  Virtual:"
.lstLanguage.AddItem "  Archivo:"
.lstLanguage.AddItem " Buscar Archivos de sonido."
.lstLanguage.AddItem " Buscar en:"
.lstLanguage.AddItem " Explorar..."
.lstLanguage.AddItem " Comenzar a Buscar"
.lstLanguage.AddItem " Detener Busqueda"
.lstLanguage.AddItem " Selecciona el directorio a buscar." '234
.lstLanguage.AddItem " Colocar caratula como Wallpaper" '235
.lstLanguage.AddItem "Opciones Biblioteca" '236
.lstLanguage.AddItem " Remover Archivos Eliminados de la Biblioteca" '237
.lstLanguage.AddItem " Agregar Audio a la Bilioteca" '238
.lstLanguage.AddItem "Recargar Base de Datos" '239
.lstLanguage.AddItem "Borrar Lista de Reproduccion" '240
.lstLanguage.AddItem "Buscar"
.lstLanguage.AddItem "Comenzar a buscar" '242
.lstLanguage.AddItem "Reproducir Seleccianados AGREGAR"
.lstLanguage.AddItem "Reproducir Todos AGREGAR" '244
.lstLanguage.AddItem "Reproducir Seleccionados NUEVO"
.lstLanguage.AddItem "Reproducir Todos NUEVO" '246
.lstLanguage.AddItem "Guardar Todo como Lista R"
.lstLanguage.AddItem "Editar Informacion de Archivos" '248
.lstLanguage.AddItem "Explorar Carpeta"
.lstLanguage.AddItem "Remover de la Biblioteca" '250
.lstLanguage.AddItem "Guardar lista de Reproduccion"
.lstLanguage.AddItem "Agregar Archivos" '252
.lstLanguage.AddItem "Quitar Tracks"
.lstLanguage.AddItem " Quitar Seleccionados" '254
.lstLanguage.AddItem " Limpiar Lista"
.lstLanguage.AddItem " Quitar Archivos Eliminados" '256
.lstLanguage.AddItem "Otras Opciones"
.lstLanguage.AddItem " Ordenar por Titulo" '258
.lstLanguage.AddItem " Ordenar por Artista"
.lstLanguage.AddItem " Ordenar por Archivo" '260
.lstLanguage.AddItem " Editar Tags Archivo(s)"
.lstLanguage.AddItem "Lista de Reproduccion" '262
.lstLanguage.AddItem " Abrir lista de Biblioteca"
.lstLanguage.AddItem " Guardar Lista de Reproduccion" '264
.lstLanguage.AddItem " Administrador de Listas"




End With
End Sub
 

Public Sub Load_Language(strLang As String)
 On Error Resume Next
 Dim Linenr As Integer
 Dim InputData
 Dim strRuta As String, strTemp As String
 With frmPopUp
  
   strRuta = tAppConfig.AppConfig & "Language\" & strLang & ".lng"
   
   If Dir(strRuta) <> "" Then
    Open strRuta For Input As #2
    .lstLanguage.Clear
    .lstLanguage.AddItem "MMPlayerX -version 3.0 language"
     Linenr = 0
     Do While Not EOF(2)
       Line Input #2, InputData
        
        If Linenr > 264 Then
          Exit Do
        End If
        If Trim(InputData) <> "" And Linenr > 0 Then
         
          If Linenr > 15 And Linenr < 28 And Linenr <> 25 Then
            strTemp = Left(LineLanguage(Linenr), 1)
            strTemp = Trim(strTemp) & "" & InputData
            .lstLanguage.AddItem Trim(strTemp)
          Else
           .lstLanguage.AddItem Trim(InputData)
          End If
        End If
        
        Linenr = Linenr + 1
     Loop
    Close #2
   Else
     Load_Language_Spanish
   End If
   ' MENU
   .mnuPlay.Caption = LineLanguage(1)
   .mnuArchivo.Caption = LineLanguage(2)
   .mnuFolder.Caption = LineLanguage(3)
   .mnuNuevaBusqueda.Caption = LineLanguage(4)
   .mnuVentanas.Caption = LineLanguage(5)
   .mnuListaR.Caption = LineLanguage(6)
   .mnuBibliotecaMultimedia.Caption = LineLanguage(7)
   .mnuEqualizador.Caption = LineLanguage(8)
   .mnuMaxSpec.Caption = LineLanguage(9)
   .mnuMCaratula.Caption = LineLanguage(10)
   .mnuTagEditor.Caption = LineLanguage(11)
   .mnuLyrics.Caption = LineLanguage(12)
   
   .mnuControles.Caption = LineLanguage(13)
   .mnuVolumen.Caption = LineLanguage(14)
   .mnuSubirVolumen.Caption = LineLanguage(15)
   .mnuBajarVolumen.Caption = LineLanguage(16)
   
   .mnuTrackAnterior.Caption = LineLanguage(17)
   frmMain.Button(0).ToolTipText = Trim(Right(LineLanguage(17), Len(LineLanguage(17)) - 1))
   frmMain.ButtonMini(0).ToolTipText = Trim(Right(LineLanguage(17), Len(LineLanguage(17)) - 1))
   
   .mnuReproducir.Caption = LineLanguage(18)
   frmMain.Button(1).ToolTipText = Trim(Right(LineLanguage(18), Len(LineLanguage(18)) - 1))
   frmMain.ButtonMini(1).ToolTipText = Trim(Right(LineLanguage(18), Len(LineLanguage(18)) - 1))
   
   .mnuPausa.Caption = LineLanguage(19)
   frmMain.Button(2).ToolTipText = Trim(Right(LineLanguage(19), Len(LineLanguage(19)) - 1))
   frmMain.ButtonMini(2).ToolTipText = Trim(Right(LineLanguage(19), Len(LineLanguage(19)) - 1))
   
   .mnuDetener.Caption = LineLanguage(20)
   frmMain.Button(3).ToolTipText = Trim(Right(LineLanguage(20), Len(LineLanguage(20)) - 1))
   frmMain.ButtonMini(3).ToolTipText = Trim(Right(LineLanguage(20), Len(LineLanguage(20)) - 1))
   
   .mnuSigTrack.Caption = LineLanguage(21)
   frmMain.Button(4).ToolTipText = Trim(Right(LineLanguage(21), Len(LineLanguage(21)) - 1))
   frmMain.ButtonMini(4).ToolTipText = Trim(Right(LineLanguage(21), Len(LineLanguage(21)) - 1))
   
   frmMain.Button(9).ToolTipText = LineLanguage(9)
   frmMain.Button(10).ToolTipText = LineLanguage(6)
   frmMain.Button(11).ToolTipText = LineLanguage(7)
   
   .mnuIntro.Caption = LineLanguage(22)
   frmMain.Button(5).ToolTipText = Trim(Right(LineLanguage(22), Len(LineLanguage(22)) - 1))
   
   .mnuSilencio.Caption = LineLanguage(23)
   frmMain.Button(6).ToolTipText = Trim(Right(LineLanguage(23), Len(LineLanguage(23)) - 1))
   
   .mnuRepetir.Caption = LineLanguage(24)
   frmMain.Button(7).ToolTipText = Trim(Right(LineLanguage(24), Len(LineLanguage(24)) - 1))
   
   .mnuOrdenAleatorio.Caption = LineLanguage(25)
   frmMain.Button(8).ToolTipText = Trim(LineLanguage(25))
   
   .mnuAtras5Seg.Caption = LineLanguage(26)
   .mnuAdelante5Seg.Caption = LineLanguage(27)
   
   .mnuOpciones.Caption = LineLanguage(28)
   .mnuSkins.Caption = LineLanguage(29)
   .mnuExpSkins.Caption = LineLanguage(30)
   .mnuWOpacity.Caption = LineLanguage(31)
   .mnuAlphaPer.Caption = LineLanguage(32)
   .mnuAcercaDe.Caption = LineLanguage(33)
   .mnuSalir.Caption = LineLanguage(34)
   
   
   frmMain.Button(12).ToolTipText = LineLanguage(39)
   frmMain.ButtonMini(5).ToolTipText = LineLanguage(39)
   frmMain.Button(13).ToolTipText = LineLanguage(40)
   frmMain.ButtonMini(6).ToolTipText = LineLanguage(40)
   frmMain.Button(14).ToolTipText = LineLanguage(41)
   frmMain.ButtonMini(7).ToolTipText = LineLanguage(41)
   frmMain.Button(15).ToolTipText = LineLanguage(42)
   frmMain.ButtonMini(8).ToolTipText = LineLanguage(42)
   
   .mnuSpecNone.Caption = LineLanguage(43)
   .mnuSpecBars.Caption = LineLanguage(44)
   .mnuSpecOsc.Caption = LineLanguage(45)
   
   
   Load_Language_Options
   load_Language_Media_Library
   load_Language_Play_List
   
   If bolCaratulaShow = True Then frmCaratula.Caption = LineLanguage(36)
   If bolAcercaShow = True Then frmAcerca.Caption = LineLanguage(35)
   If bolLyricsShow = True Then frmLyrics.Caption = LineLanguage(37): frmLyrics.lblNoLyrics.Caption = LineLanguage(38)
   If bolTagsShow = True Then Load_Language_Tags
   If bolSearchShow = True Then Load_Language_Search
   
   
   '//change language at systray icons
   If PlayerTrayIcon.Previous = True Then CambiarIcono frmMain.txtSTIcon(0).hwnd, frmMain.ImageList.ListImages(1).ExtractIcon.Handle, frmMain.Button(0).ToolTipText & " - MMPlayerX"
   If PlayerTrayIcon.Play = True Then CambiarIcono frmMain.txtSTIcon(1).hwnd, frmMain.ImageList.ListImages(2).ExtractIcon.Handle, frmMain.Button(1).ToolTipText & " - MMPlayerX"
   If PlayerTrayIcon.Pause = True Then CambiarIcono frmMain.txtSTIcon(2).hwnd, frmMain.ImageList.ListImages(3).ExtractIcon.Handle, frmMain.Button(2).ToolTipText & " - MMPlayerX"
   If PlayerTrayIcon.Stop = True Then CambiarIcono frmMain.txtSTIcon(3).hwnd, frmMain.ImageList.ListImages(4).ExtractIcon.Handle, frmMain.Button(3).ToolTipText & " - MMPlayerX"
   If PlayerTrayIcon.Next = True Then CambiarIcono frmMain.txtSTIcon(4).hwnd, frmMain.ImageList.ListImages(5).ExtractIcon.Handle, frmMain.Button(4).ToolTipText & " - MMPlayerX"
  
 End With
End Sub

Sub load_Language_Play_List()
With frmPlayList
    .BTN(0).ToolTipText = LineLanguage(252)
    .BTN(1).ToolTipText = LineLanguage(253)
    .BTN(2).ToolTipText = LineLanguage(257)
    .BTN(3).ToolTipText = LineLanguage(262)
End With

With frmPopUp
    .mnuQuitarS.Caption = LineLanguage(254)
    .mnuLimpiarL.Caption = LineLanguage(255)
    .mnuQuitarAB.Caption = LineLanguage(256)
    
    .mnuOrdenarTitulo.Caption = LineLanguage(258)
    .mnuOrdenarArtista.Caption = LineLanguage(259)
    .mnuOrdenarArchivo.Caption = LineLanguage(260)
    .mnuEditarArchivos.Caption = LineLanguage(261)
    
    .mnuAbrirL.Caption = LineLanguage(263)
    .mnuGuardarL.Caption = LineLanguage(264)
    .mnuAdministrador.Caption = LineLanguage(265)
    
End With

End Sub

Sub load_Language_Media_Library()
With frmLibrary
    .BTN(0).ToolTipText = LineLanguage(236)
    frmPopUp.mnuRemoverbiblioteca.Caption = LineLanguage(237)
    frmPopUp.mnuAgregarAudio.Caption = LineLanguage(238)
    .BTN(1).ToolTipText = LineLanguage(239)
    .BTN(2).ToolTipText = LineLanguage(240)
    .Label2.Caption = LineLanguage(241)
    .BTN(3).ToolTipText = LineLanguage(242)
End With

With frmPopUp
   .mnuReproducirSeleccionadosAgregar.Caption = LineLanguage(243)
   .mnuReproducirTodosAgregar.Caption = LineLanguage(244)
   .mnuReproducirSeleccionadosNuevo.Caption = LineLanguage(245)
   .mnuReproducirTodosNuevo.Caption = LineLanguage(246)
   .mnuGuardarTodoLista.Caption = LineLanguage(247)
   .mnuEditarInformacionArchivos.Caption = LineLanguage(248)
   .mnuExplorarCarpeta.Caption = LineLanguage(249)
   .mnuRemoverbibliotecaM.Caption = LineLanguage(250)
End With

End Sub

Sub Load_Language_Search()
  With frmSearch
      .Caption = LineLanguage(229)
      .Label.Caption = LineLanguage(230)
      .cmdBrowse.Caption = LineLanguage(231)
      
    If bSearching = False Then
      .cmdSearch.Caption = LineLanguage(232)
    Else
      .cmdSearch.Caption = LineLanguage(233)
    End If
 End With
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Public Sub Load_Language_Options()
 On Error Resume Next
 Dim i As Integer
 With frmOpciones
   .Caption = LineLanguage(70)
   
   '// APLICACION
   .TreeOptions.Nodes("Application").Text = LineLanguage(74)
    .TSAppConfig.Tabs(1).Caption = LineLanguage(75)
     .lblApp(3).Caption = LineLanguage(76)
     .cmdAppConfig.Caption = LineLanguage(77)
     .lblApp(6).Caption = LineLanguage(78)
     .lblApp(7).Caption = LineLanguage(79)
    .TSAppConfig.Tabs(2).Caption = LineLanguage(80)
     .lblApp(0).Caption = LineLanguage(81)
     .chkWindowsState(2).Caption = LineLanguage(82)
     .chkWindowsState(3).Caption = LineLanguage(83)
     .chkWindowsState(4).Caption = LineLanguage(84)
     .chkDir.Caption = LineLanguage(85)
     .lblApp(1).Caption = LineLanguage(86)
     .chkWindowsState(0).Caption = LineLanguage(87)
     .chkWindowsState(1).Caption = LineLanguage(88)
    .TSAppConfig.Tabs(3).Caption = LineLanguage(89)
     .lblApp(2).Caption = LineLanguage(90)
     
  '// SKINS
    .TreeOptions.Nodes("Skins").Text = LineLanguage(91)
     .lblSkin(1).Caption = LineLanguage(92)
     .lblSkin(0).Caption = LineLanguage(93)
     .chkUseFile.Caption = LineLanguage(94)
     
  '// WALLPAPER
    .TreeOptions.Nodes("Wallpaper").Text = LineLanguage(95)
     .chkWallpaper.Caption = LineLanguage(235)
     .lblWallpaper.Caption = LineLanguage(9)
     .optWallpaper(0).Caption = LineLanguage(97)
     .optWallpaper(3).Caption = LineLanguage(98)
     .optWallpaper(2).Caption = LineLanguage(99)
     .optWallpaper(1).Caption = LineLanguage(100)
     .chkProporcional.Caption = LineLanguage(101)
     
  '// PLAY LIST FORMAT
    .TreeOptions.Nodes("ScrollText").Text = LineLanguage(102)
     .lblPL(0).Caption = LineLanguage(103)
     .lblPL(1).Caption = LineLanguage(104)
     .lblPL(2).Caption = LineLanguage(105)
     .optScrollType(0).Caption = LineLanguage(106)
     .optScrollType(1).Caption = LineLanguage(107)
     .lblPL(3).Caption = LineLanguage(108)
     
  '// REPRODUCTOR
    .TreeOptions.Nodes("Player").Text = LineLanguage(109)
     .lblPlayer(0).Caption = LineLanguage(110)
     .lblPlayer(1).Caption = LineLanguage(111)
     .chkPIcon(0).Caption = LineLanguage(112)
     .chkPIcon(1).Caption = LineLanguage(113)
     .chkPIcon(2).Caption = LineLanguage(114)
     .chkPIcon(3).Caption = LineLanguage(115)
     .chkPIcon(4).Caption = LineLanguage(116)
     .lblPlayer(2).Caption = LineLanguage(117)
     .lblPlayer(3).Caption = LineLanguage(118)
     .chkPlayStart.Caption = LineLanguage(119)
     
  '// EFECTOS DSP FX
    .TreeOptions.Nodes("Effects").Text = LineLanguage(120)
     .tsDSP.Tabs(1).Caption = LineLanguage(121)
     .chkDSP(0).Caption = LineLanguage(122)
      .lblChorus(0).Caption = LineLanguage(123)
      .lblChorus(1).Caption = LineLanguage(124)
      .lblChorus(2).Caption = LineLanguage(125)
      .lblChorus(3).Caption = LineLanguage(126)
      .lblChorus(4).Caption = LineLanguage(127)
      .lblChorus(5).Caption = LineLanguage(128)
      .lblChorus(6).Caption = LineLanguage(129)
      
     .tsDSP.Tabs(2).Caption = LineLanguage(130)
     .chkDSP(1).Caption = LineLanguage(131)
      .lblComp(0).Caption = LineLanguage(132)
      .lblComp(1).Caption = LineLanguage(133)
      .lblComp(2).Caption = LineLanguage(134)
      .lblComp(3).Caption = LineLanguage(135)
      .lblComp(4).Caption = LineLanguage(136)
      .lblComp(5).Caption = LineLanguage(137)
      
     .tsDSP.Tabs(3).Caption = LineLanguage(138)
     .chkDSP(2).Caption = LineLanguage(139)
      .lblDis(0).Caption = LineLanguage(140)
      .lblDis(1).Caption = LineLanguage(141)
      .lblDis(2).Caption = LineLanguage(142)
      .lblDis(3).Caption = LineLanguage(143)
      .lblDis(4).Caption = LineLanguage(144)
      
     .tsDSP.Tabs(4).Caption = LineLanguage(145)
     .chkDSP(3).Caption = LineLanguage(146)
      .lblEcho(0).Caption = LineLanguage(147)
      .lblEcho(1).Caption = LineLanguage(148)
      .lblEcho(2).Caption = LineLanguage(149)
      .lblEcho(3).Caption = LineLanguage(150)
      .lblEcho(4).Caption = LineLanguage(151)
      
     .tsDSP.Tabs(5).Caption = LineLanguage(152)
     .chkDSP(4).Caption = LineLanguage(153)
      .lblFlan(0).Caption = LineLanguage(154)
      .lblFlan(1).Caption = LineLanguage(155)
      .lblFlan(2).Caption = LineLanguage(156)
      .lblFlan(3).Caption = LineLanguage(157)
      .lblFlan(4).Caption = LineLanguage(158)
      .lblFlan(5).Caption = LineLanguage(159)
      .lblFlan(6).Caption = LineLanguage(160)
      
     .tsDSP.Tabs(6).Caption = LineLanguage(161)
     .chkDSP(5).Caption = LineLanguage(162)
      .lblGarg(0).Caption = LineLanguage(163)
      .lblGarg(1).Caption = LineLanguage(164)
      
     .tsDSP.Tabs(7).Caption = LineLanguage(165)
     .chkDSP(6).Caption = LineLanguage(166)
      .lblL2(0).Caption = LineLanguage(167)
      .lblL2(1).Caption = LineLanguage(168)
      .lblL2(2).Caption = LineLanguage(169)
      .lblL2(3).Caption = LineLanguage(170)
      .lblL2(4).Caption = LineLanguage(171)
      .lblL2(5).Caption = LineLanguage(172)
      .lblL2(6).Caption = LineLanguage(173)
      .lblL2(7).Caption = LineLanguage(174)
      .lblL2(8).Caption = LineLanguage(175)
      .lblL2(9).Caption = LineLanguage(176)
      .lblL2(10).Caption = LineLanguage(177)
      .lblL2(11).Caption = LineLanguage(178)
      
     .tsDSP.Tabs(8).Caption = LineLanguage(179)
     .chkDSP(8).Caption = LineLanguage(180)
      .lblWaves(0).Caption = LineLanguage(181)
      .lblWaves(1).Caption = LineLanguage(182)
      .lblWaves(2).Caption = LineLanguage(183)
      .lblWaves(3).Caption = LineLanguage(184)
     
     .cmdDSPReset.Caption = LineLanguage(185)
     .cmdDSPClear.Caption = LineLanguage(186)
     
  '// EQUALIZADOR
    .TreeOptions.Nodes("Equalizer").Text = LineLanguage(187)
     .chkDSP(7).Caption = LineLanguage(188)
     .lblEQ(10).Caption = LineLanguage(189)
     .cmdDeleteEQ.Caption = LineLanguage(190)
     .cmdSaveEQ.Caption = LineLanguage(191)
     
  '// VISUALISATION
    .TreeOptions.Nodes("Visualization").Text = LineLanguage(194)
     .lblCurrentVis(0) = LineLanguage(195)
     .lblCurrentVis(1) = LineLanguage(196)
     .lblCurrentVis(2) = LineLanguage(197)
     
     For i = 198 To 214
      .lblVis(i - 198).Caption = LineLanguage(i)
     Next i
     
     .cmdVisualizacion(0).Caption = LineLanguage(215)
     .cmdVisualizacion(4).Caption = LineLanguage(216)
     .cmdVisualizacion(1).Caption = LineLanguage(217)
     .cmdVisualizacion(2).Caption = LineLanguage(218)
     
     frmPopUp.mnuPrevVis.Caption = LineLanguage(221)
     frmPopUp.mnuNextVis.Caption = LineLanguage(222)
     frmPopUp.mnuConfigVis.Caption = LineLanguage(223)
     frmPopUp.mnuExit.Caption = LineLanguage(224)
     
     .lblApp(8).Caption = LineLanguage(226)
     .lblApp(9).Caption = LineLanguage(227)
     .lblApp(10).Caption = LineLanguage(228)
    
   '//buttons
   .cmdOk.Caption = LineLanguage(71)
   .cmdCancel.Caption = LineLanguage(72)
   .cmdApply.Caption = LineLanguage(73)
   .cmdSaveConfig.Caption = LineLanguage(225)
  End With
End Sub

Sub Load_Language_Tags()
 With frmTags
    .Caption = LineLanguage(57)
    .cmdOk.Caption = LineLanguage(64)
    .cmdCancel.Caption = LineLanguage(65)
    .cmdApply.Caption = LineLanguage(66)
    .cmdSelAll.Caption = LineLanguage(59)
    .TabStrip.Tabs(1).Caption = LineLanguage(60)
    .TabStrip.Tabs(3).Caption = LineLanguage(61)
    .TabStrip.Tabs(2).Caption = LineLanguage(67)
    .cmdAdd.Caption = LineLanguage(62)
    .cmdUndo.Caption = LineLanguage(63)
    .cmdAddArt.Caption = LineLanguage(68)
    .cmdRemoveArt.Caption = LineLanguage(69)
 End With
End Sub

Public Function LineLanguage(Number As Integer) As String
 On Error Resume Next
  
   If frmPopUp.lstLanguage.ListCount = 0 Then Exit Function
   If Number > frmPopUp.lstLanguage.ListCount - 1 Then Exit Function
   LineLanguage = Trim(frmPopUp.lstLanguage.List(Number))
  
End Function

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

