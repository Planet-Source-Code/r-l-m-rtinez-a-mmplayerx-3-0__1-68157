VERSION 5.00
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Begin VB.Form frmPlayList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   248
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClientArea 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5385
      Left            =   0
      ScaleHeight     =   359
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   3
      Top             =   0
      Width           =   3465
      Begin VSFlex6Ctl.vsFlexGrid FGS 
         Bindings        =   "frmPlayList.frx":0000
         DragIcon        =   "frmPlayList.frx":0014
         Height          =   1170
         Left            =   1080
         TabIndex        =   4
         Top             =   3630
         Visible         =   0   'False
         Width           =   1395
         _cx             =   2461
         _cy             =   2064
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         MultiTotals     =   0   'False
         SubtotalPosition=   0
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
         DataMember      =   ""
      End
      Begin MMPlayerXProject.Slider Slider 
         Height          =   2010
         Left            =   3150
         TabIndex        =   5
         Top             =   1410
         Visible         =   0   'False
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   3545
         BackColor       =   65535
         Value           =   100
      End
      Begin VSFlex6Ctl.vsFlexGrid PL 
         Bindings        =   "frmPlayList.frx":031E
         DragIcon        =   "frmPlayList.frx":0332
         Height          =   2325
         Left            =   735
         TabIndex        =   6
         Top             =   0
         Width           =   2445
         _cx             =   4313
         _cy             =   4101
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   12632256
         ForeColor       =   16777215
         BackColorFixed  =   12632256
         ForeColorFixed  =   65280
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483634
         BackColorBkg    =   12632256
         BackColorAlternate=   12632256
         GridColor       =   0
         GridColorFixed  =   4210752
         TreeColor       =   255
         FloodColor      =   192
         SheetBorder     =   65280
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         MultiTotals     =   0   'False
         SubtotalPosition=   0
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
         DataMember      =   ""
      End
      Begin VSFlex6Ctl.vsFlexGrid PL2 
         Bindings        =   "frmPlayList.frx":063C
         DragIcon        =   "frmPlayList.frx":0650
         Height          =   1095
         Left            =   75
         TabIndex        =   7
         Top             =   3600
         Visible         =   0   'False
         Width           =   1950
         _cx             =   3440
         _cy             =   1931
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   0
         ForeColor       =   16777215
         BackColorFixed  =   0
         ForeColorFixed  =   65280
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483634
         BackColorBkg    =   0
         BackColorAlternate=   4210752
         GridColor       =   0
         GridColorFixed  =   4210752
         TreeColor       =   255
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   8
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         MultiTotals     =   0   'False
         SubtotalPosition=   0
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0   'False
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
         DataMember      =   ""
      End
   End
   Begin MMPlayerXProject.Button BTN 
      Height          =   255
      Index           =   1
      Left            =   525
      TabIndex        =   0
      ToolTipText     =   "Quitar"
      Top             =   5775
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   450
      ButtonColor     =   255
      MousePointer    =   1
      Style           =   1
   End
   Begin MMPlayerXProject.Button BTN 
      Height          =   255
      Index           =   2
      Left            =   990
      TabIndex        =   1
      ToolTipText     =   "Organizar"
      Top             =   5835
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   450
      ButtonColor     =   255
      MousePointer    =   1
      Style           =   1
   End
   Begin MMPlayerXProject.Button BTN 
      Height          =   255
      Index           =   3
      Left            =   1455
      TabIndex        =   2
      ToolTipText     =   "Listas"
      Top             =   5835
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   450
      ButtonColor     =   255
      MousePointer    =   1
      Style           =   1
   End
   Begin MMPlayerXProject.Button BTN 
      Height          =   255
      Index           =   0
      Left            =   15
      TabIndex        =   8
      ToolTipText     =   "Quitar"
      Top             =   5820
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   450
      ButtonColor     =   255
      MousePointer    =   1
      Style           =   1
   End
End
Attribute VB_Name = "frmPlayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cWindows As New cWindowSkin
Dim cAjustarDesk As New clsDockingHandler
Dim InFormDrag As Boolean
Dim bMakingListRep        As Boolean

Dim iContFilas As Integer
Dim bCtrlSel As Boolean
Dim bShiftSel As Boolean
Dim iSelRow As Integer
Dim i As Integer
Dim j As Integer
Dim iButtonsLeft As Integer, iButtonsTop As Integer
Dim bDblClkPlay As Boolean

Sub Remover_Archivos_Eliminados()
On Error Resume Next
Dim X As Integer

For i = 0 To PL.Rows - 1
  
 If Dir(PL.TextMatrix(i, 10)) = "" Then
    
    PL.RemoveItem i - X
    X = X + 1
 End If
 PL.TextMatrix(i, 0) = i + 1 & "."
Next i
End Sub

Sub Remover_Tracks()
On Error Resume Next
Dim X As Integer
If PL.Rows = 0 Then Exit Sub
If FGS.Row = -1 Then
    PL.RemoveItem (PL.Row)
    For X = 0 To PL.Rows - 1
        PL.TextMatrix(X, 0) = X + 1 & "."
    Next X
    Exit Sub
End If
FGS.Col = 0
FGS.Sort = flexSortNumericAscending

For i = 0 To FGS.Rows - 1
    PL.RemoveItem CInt(FGS.TextMatrix(i, 0)) - X
    X = X + 1
Next

For X = 0 To PL.Rows - 1
    PL.TextMatrix(X, 0) = X + 1 & "."
Next X

PL.Select CInt(FGS.TextMatrix(i - 1, 0)), 0
FGS.Rows = 0
End Sub


Sub Editar_Archivos()
On Error Resume Next
If PL.Rows = 0 Then Exit Sub
If FGS.Row = -1 Then
    frmTags.listRef.ListItems.Clear
    frmTags.fileTags.Clear
    frmTags.Load_Tags PL.TextMatrix(CInt(PL.Row), PL.Cols - 1), CInt(PL.Row)
    frmTags.Show
    Exit Sub
End If

    frmTags.listRef.ListItems.Clear
    frmTags.fileTags.Clear
For i = 0 To FGS.Rows - 1
    frmTags.Load_Tags PL.TextMatrix(CInt(FGS.TextMatrix(i, 0)), PL.Cols - 1), CInt(FGS.TextMatrix(i, 0))
Next
    frmTags.Show
End Sub

Sub Agregar_PlayList_de_BaseDatos(iRow As Integer)
 On Error Resume Next
 Dim sFileName As String, sFileEx As String
 Dim sFormat As String, sNewString As String, SplitField() As String, CleanStr As String
 Dim i As Integer
 Dim iIndex As Integer, iSpaces As Integer
 Dim aFile() As String
   
   bMakingListRep = True
   
   aFile = Split(PL.TextMatrix(iRow, PL.Cols - 1), "\", , vbTextCompare)
   sFileName = Left(aFile(UBound(aFile)), Len(aFile(UBound(aFile))) - 4)
   sFileEx = Right(PL.TextMatrix(iRow, PL.Cols - 1), 3)
       
   sFormat = ""
   '// Song Name
   sFormat = Replace(sFormatPlayList, "%S", PL.TextMatrix(iRow, 2))
   '// Artist
   sFormat = Replace(sFormat, "%A", PL.TextMatrix(iRow, 3))
   '// Album
   sFormat = Replace(sFormat, "%B", PL.TextMatrix(iRow, 4))
   '// Year
   sFormat = Replace(sFormat, "%Y", PL.TextMatrix(iRow, 6))
   '// Genre
   sFormat = Replace(sFormat, "%G", PL.TextMatrix(iRow, 5))
   '// Time
   sFormat = Replace(sFormat, "%T", PL.TextMatrix(iRow, 7))
   '// File Name
   sFormat = Replace(sFormat, "%N", sFileName)
   '// Ful Path
   sFormat = Replace(sFormat, "%P", PL.TextMatrix(iRow, PL.Cols - 1))
   '// File extencion
   sFormat = Replace(sFormat, "%F", sFileEx)
   
   If sFormat = sFormatPlayList Then sFormat = sFileName
      
   '------------------------------------------------------------------------------
    CleanStr = Trim$(sFormat)
    
    'Upper case and / or lower case the string correctly.
    SplitField = Split(CleanStr, " ", , vbTextCompare)
    CleanStr = ""
    For iSpaces = 0 To UBound(SplitField)
        If Not iSpaces = 0 Or Not IsNumeric(SplitField(iSpaces)) Then
          sNewString = UCase$(Left$(SplitField(iSpaces), 1))
          sNewString = sNewString & LCase$(Right$(SplitField(iSpaces), Len(SplitField(iSpaces)) - 1))
          CleanStr = CleanStr & sNewString & " "
        End If
    Next iSpaces
    sFormat = Trim$(CleanStr)
  '------------------------------------------------------------------------------
  PL.TextMatrix(iRow, 1) = sFormat
  PL.TextMatrix(iRow, 0) = iRow + 1 & "."
  Show_ScrollBar
 
bMakingListRep = False

End Sub



Sub Show_ScrollBar()
  On Error Resume Next
  Dim sngHeight As Single
  Dim iRows As Integer
  '// calculate the height for any text
  sngHeight = PL.RowHeight(0) / Screen.TwipsPerPixelY
  
  '// you not always show the scroll bar only in case requiered
  If (sngHeight * PL.Rows) > PL.Height Then
    Slider.Visible = True
    iRows = PL.Height / sngHeight
    
    Slider.min = 0
    Slider.Max = PL.Rows - CInt(iRows)
    Slider.Value = PL.Rows - CInt(iRows)
  Else
    Slider.Visible = False
  End If

End Sub



Private Sub BTN_Click(Index As Integer)
Select Case Index
  Case 0
    PopupMenu frmPopUp.mnuPlay
  Case 1
    PopupMenu frmPopUp.mnuQuitar
  Case 2
    PopupMenu frmPopUp.mnuMis
  Case 3
    PopupMenu frmPopUp.mnuLista
End Select

End Sub
Sub Guardar_PlayList()
 Dim i As Integer
 Dim j As Integer
 Dim archivoINI As String
 Dim s As String
 On Error GoTo BITCH

PL2.Cols = 9
PL2.Rows = PL.Rows + 1
For i = 0 To PL.Rows - 1
    For j = 0 To PL.Cols - 3
        PL2.TextMatrix(i + 1, j) = PL.TextMatrix(i, j + 2)
    Next j
Next i
 
 If Dir(tAppConfig.AppConfig & "Library\", vbDirectory) = "" Then Exit Sub
 
 s = InputBox(LineLanguage(220), "Play List", "PlayList_X", Me.Left + (Me.Width / 2) - 2000, Me.Top + 2000)
 If Trim(s) = "" Then Exit Sub
 archivoINI = tAppConfig.AppConfig & "Library\" & s & ".pls"

If Dir(archivoINI) <> "" Then '// si existe el archivo borrarlo
 SetAttr archivoINI, vbNormal
 Kill archivoINI
End If
    
PL2.SaveGrid archivoINI, flexFileAll
frmPopUp.Cargar_Lista_Rep
frmLibrary.TreeFiles.Nodes.Add "kPlaL", tvwChild, "kPlaE" & s, s, 14


Exit Sub
BITCH:
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 Then
    If bCtrlSel = False Then
    bCtrlSel = True
    iContFilas = 0
    FGS.Rows = 1
    FGS.Col = 0
    FGS.Rows = iContFilas + 1
    FGS.Row = iContFilas
    FGS.Text = PL.Row
    iContFilas = iContFilas + 1
    PL.Col = 0
    PL.CellBackColor = PL.BackColorSel
    PL.CellForeColor = PL.ForeColorSel
    PL.Col = 1
    PL.CellBackColor = PL.BackColorSel
    PL.CellForeColor = PL.ForeColorSel
    PL.Col = 7
    PL.CellBackColor = PL.BackColorSel
    PL.CellForeColor = PL.ForeColorSel
    End If
End If

End Sub

Private Sub Form_Load()
    
  bolPlayListShow = True
  Set cWindows.FormularioPadre = Me
  Set cAjustarDesk.ParentForm = Me

End Sub

Sub cargar_formulario()
Dim iX As Integer, iY As Integer

  cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
  iX = Read_INI("Configuration", "ExitButtonX", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
  iY = Read_INI("Configuration", "ExitButtonY", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
  cWindows.ButtonExitXY CLng(iX), CLng(iY)
  cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
  cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
  iButtonsLeft = Read_INI("Configuration", "ButtonsLeft", 5, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
  iButtonsTop = Read_INI("Configuration", "ButtonsTop", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
  cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\"
  PL.Cols = 11
  PL.ColHidden(0) = False
  PL.ColHidden(1) = False
  PL.ColHidden(2) = True
  PL.ColHidden(3) = True
  PL.ColHidden(4) = True
  PL.ColHidden(5) = True
  PL.ColHidden(6) = True
  PL.ColHidden(7) = False
  PL.ColHidden(8) = True
  PL.ColHidden(9) = True
  PL.ColHidden(10) = True
  PL.ColWidth(0) = 500
  PL.ColWidth(1) = 3000
  PL.ColWidth(7) = 500
  PL.ColAlignment(7) = flexAlignRightCenter
    
'  PL.CellAlignment = flexAlignRightCenter
  
    picClientArea.Left = cWindows.AreaLeft
    picClientArea.Top = cWindows.AreaTop
    picClientArea.Width = cWindows.AreaWidth
    picClientArea.Height = cWindows.AreaHeight
    PL.Left = 0
    PL.Height = cWindows.AreaHeight

    Slider.Height = cWindows.AreaHeight
    BTN(0).Left = 30
    BTN(0).Top = Me.ScaleHeight - BTN(0).Height - 10
    BTN(1).Top = Me.ScaleHeight - BTN(1).Height - 10
    BTN(1).Left = BTN(0).Left + BTN(0).Width + 1
    BTN(2).Top = Me.ScaleHeight - BTN(2).Height - 10
    BTN(2).Left = BTN(1).Left + BTN(1).Width + 1
    BTN(3).Top = Me.ScaleHeight - BTN(3).Height - 10
    BTN(3).Left = BTN(2).Left + BTN(2).Width + 1

    Show_ScrollBar
    If Slider.Visible = True Then
        PL.Width = cWindows.AreaWidth + 10
       PL.ColWidth(1) = (picClientArea.Width * 15) - PL.ColWidth(7) - (Slider.Width * 15) - PL.ColWidth(0)
       Slider.Left = picClientArea.Width - Slider.Width
    Else
       PL.Width = cWindows.AreaWidth
       PL.ColWidth(1) = (picClientArea.Width * 15) - PL.ColWidth(7) - PL.ColWidth(0)
       Slider.Left = picClientArea.Width - Slider.Width
    End If
    
    Slider.Top = 0

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
cWindows.Formulario_Down X, Y
cAjustarDesk.StartDockDrag X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
InFormDrag = True

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 cWindows.Formulario_MouseMove Button, X, Y

 If cWindows.Ajustando = True Then
    picClientArea.Left = cWindows.AreaLeft
    picClientArea.Top = cWindows.AreaTop
    picClientArea.Width = cWindows.AreaWidth
    picClientArea.Height = cWindows.AreaHeight
    PL.Height = cWindows.AreaHeight
    BTN(0).Left = iButtonsLeft
    BTN(0).Top = Me.ScaleHeight - BTN(0).Height - iButtonsTop
    BTN(1).Top = Me.ScaleHeight - BTN(1).Height - iButtonsTop
    BTN(1).Left = BTN(0).Left + BTN(0).Width + 1
    BTN(2).Top = Me.ScaleHeight - BTN(2).Height - iButtonsTop
    BTN(2).Left = BTN(1).Left + BTN(1).Width + 1
    BTN(3).Top = Me.ScaleHeight - BTN(3).Height - iButtonsTop
    BTN(3).Left = BTN(2).Left + BTN(2).Width + 1
    Slider.Top = 0
    Slider.Height = PL.Height
    Show_ScrollBar
    If Slider.Visible = True Then
'       PL.Width = cWindows.AreaWidth - Slider.Width
        PL.Width = cWindows.AreaWidth + 10
       PL.ColWidth(1) = (picClientArea.Width * 15) - PL.ColWidth(7) - (Slider.Width * 15) - PL.ColWidth(0)
       Slider.Left = picClientArea.Width - Slider.Width
    Else
       PL.Width = cWindows.AreaWidth
       PL.ColWidth(1) = (picClientArea.Width * 15) - PL.ColWidth(7) - PL.ColWidth(0)
       Slider.Left = picClientArea.Width - Slider.Width
    End If

    
 End If
    If InFormDrag And cWindows.Ajustando = False Then
        ' Continue window draggin'
        cAjustarDesk.UpdateDockDrag X * Screen.TwipsPerPixelX, _
            Y * Screen.TwipsPerPixelY
         bHookForm = False
        Exit Sub
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If cWindows.Ajustando = True Then Show_ScrollBar
cWindows.Formulario_MouseUp X, Y
If cWindows.ClickExitButton = True Then
    frmMain.Mostrar_Play_List
End If

InFormDrag = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 1
   Set cWindows = Nothing
   Set cAjustarDesk = Nothing
End Sub

Sub Next_Track()
    
    On Error Resume Next
    
    If iIndexPlay + 1 > (PL.Rows - 1) Then
       Play_Track 0
    Else
       Play_Track iIndexPlay + 1
    End If
    
End Sub

Sub Previous_Track()
    
    On Error Resume Next
    
    If iIndexPlay - 1 < 0 Then
       Play_Track PL.Rows - 1
    Else
       Play_Track iIndexPlay - 1
    End If

End Sub

Sub Play_Track(iIndex As Integer)
    Dim iAnt As Integer
    Dim i As Integer
    Dim color As Long
    On Error Resume Next
   PL.HighLight = flexHighlightNever
    For i = 0 To PL.Cols - 1
    
         PL.Row = iIndexPlay
         PL.Col = i
         
         If (iIndexPlay + 1) Mod 2 <> 0 Then
            color = PL.BackColor
         Else
            color = PL.BackColorAlternate
         End If

         PL.CellBackColor = color
         PL.CellForeColor = PL.ForeColor
         
         PL.Row = iIndex
         PL.CellBackColor = BackColorPlaying
         PL.CellForeColor = ForeColorPlaying
    Next
   PL.HighLight = flexHighlightAlways
 iIndexPlay = iIndex
 sFileMainPlaying = PL.TextMatrix(iIndexPlay, 10)
 
 If bLoading = True Then Exit Sub
 If PL.Rows = 0 Then Exit Sub
 
 If Dir(sFileMainPlaying) = "" Then Exit Sub
 frmMain.PlayerIsPlaying = "true"
 frmMain.Play
 frmLibrary.Actualizar_Track sFileMainPlaying
 '/kolokar la lista en laposicion adecuada
  Dim sngHeight As Single
  Dim iRows As Integer

If bDblClkPlay = False Then
  If Slider.Visible = True Then
    '// calculate the height for any text
    sngHeight = PL.RowHeight(0) / Screen.TwipsPerPixelY
    iRows = (PL.Height / sngHeight) / 2
    If iIndexPlay - iRows >= 0 Then
    PL.TopRow = iIndexPlay - iRows
    Else
    PL.TopRow = 0
    End If
  End If
End If
 bDblClkPlay = False
 End Sub


Private Sub PL_DblClick()
 On Error Resume Next
 bDblClkPlay = True
 Play_Track PL.Row
End Sub

Private Sub PL_GotFocus()
bFocusPlayList = True
End Sub

Private Sub PL_LostFocus()
bFocusPlayList = False
End Sub

Private Sub Slider_GotFocus()
bFocusPlayList = True
End Sub

Private Sub Slider_LostFocus()
bFocusPlayList = False
End Sub


Private Sub PL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 2 Then
    bCtrlSel = True
Else
    bCtrlSel = False
End If

End Sub

Private Sub PL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbDefault
End Sub


Private Sub picClientArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Paint()
'If bLoading = True Then Exit Sub
cWindows.Formulario_Paint
End Sub

Sub Agregar_PlayList_de_Archivo(sFullPath As String)
 On Error Resume Next
 Dim clsMp3 As New cMP3
 Dim iRow As Integer
 If Dir(sFullPath) = "" Then Exit Sub
    
'"TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"
   ' load tags
   clsMp3.Read_MPEGInfo = True
   clsMp3.Read_File_Tags sFullPath
   
   PL.Rows = PL.Rows + 1
   iRow = PL.Rows - 1
   PL.TextMatrix(iRow, 2) = clsMp3.Title
   PL.TextMatrix(iRow, 3) = clsMp3.Artist
   PL.TextMatrix(iRow, 4) = clsMp3.Album
   PL.TextMatrix(iRow, 5) = clsMp3.Genre
   PL.TextMatrix(iRow, 6) = clsMp3.Year
   PL.TextMatrix(iRow, 7) = clsMp3.MPEG_DurationTime
   PL.TextMatrix(iRow, 8) = 0
   PL.TextMatrix(iRow, 9) = 0
   PL.TextMatrix(iRow, 10) = sFullPath
   Agregar_PlayList_de_BaseDatos iRow
   
 Set clsMp3 = Nothing
   
End Sub



Private Sub PL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo HELL
Dim color As Long
Dim iFila As Integer

If bCtrlSel = False And FGS.Rows = 0 Then
 Dim NRows As Integer, iStart As Integer, iEnd As Integer, iTemp As Integer
  NRows = Abs(PL.RowSel - PL.Row)
  If NRows = 0 Then Exit Sub
  iStart = PL.Row
  iEnd = PL.RowSel
  FGS.Rows = NRows + 1
  If iStart > iEnd Then '// ajustar los valosres si es una selccion invertida
    iTemp = iStart
    iStart = iEnd
    iEnd = iTemp
  End If
  
  j = 0
  For i = iStart To iEnd
     FGS.Row = j
     FGS.Text = i
     j = j + 1
  Next i
 bShiftSel = True
 Exit Sub
End If


If bCtrlSel = True Then
    
    For i = 0 To FGS.Rows - 1
       FGS.Col = 0
       FGS.Row = i
       If PL.Text = FGS.Text Then Exit Sub
     Next i
    
    FGS.Col = 0
    FGS.Rows = iContFilas + 1
    FGS.Row = iContFilas
    FGS.Text = PL.Row
    iContFilas = iContFilas + 1
    
    '// Seleccionar fila
    
    PL.Col = 0
    PL.CellBackColor = PL.BackColorSel
    PL.CellForeColor = PL.ForeColorSel
    PL.Col = 1
    PL.CellBackColor = PL.BackColorSel
    PL.CellForeColor = PL.ForeColorSel
    PL.Col = 7
    PL.CellBackColor = PL.BackColorSel
    PL.CellForeColor = PL.ForeColorSel
    iFila = PL.Row
    Exit Sub
End If




If bCtrlSel = False And FGS.Rows > 0 And bShiftSel = False Then
  iFila = PL.Row
   For i = 0 To FGS.Rows - 1
      FGS.Row = i
      
      If CInt(FGS.Text) Mod 2 <> 0 Then
        color = PL.BackColor
      Else
        color = PL.BackColorAlternate
      End If
      
      If iIndexPlay <> CInt(FGS.Text) Then
        PL.Row = CInt(FGS.Text)
        PL.Col = 0
        PL.CellBackColor = color
        PL.CellForeColor = PL.ForeColor
        PL.Col = 1
        PL.CellBackColor = color
        PL.CellForeColor = PL.ForeColor
        PL.Col = 7
        PL.CellBackColor = color
        PL.CellForeColor = PL.ForeColor
      Else
        PL.Row = CInt(FGS.Text)
        PL.Col = 0
        PL.CellBackColor = BackColorPlaying
        PL.CellForeColor = ForeColorPlaying
        PL.Col = 1
        PL.CellBackColor = BackColorPlaying
        PL.CellForeColor = ForeColorPlaying
        PL.Col = 7
        PL.CellBackColor = BackColorPlaying
        PL.CellForeColor = ForeColorPlaying
      End If
    Next i
    FGS.Rows = 0
    PL.Row = iFila
    bCtrlSel = False
    Exit Sub
End If

bShiftSel = False
FGS.Rows = 0
Exit Sub
HELL:

End Sub

Private Sub PL_Scroll()
 Slider.Value = Slider.Max - PL.TopRow

End Sub

Sub Slider_Change(Value As Long)
On Error Resume Next
Dim iRow As Integer
iRow = Slider.Max - CInt(Slider.Value)
If iRow < 0 Then iRow = 0
PL.TopRow = iRow

End Sub


