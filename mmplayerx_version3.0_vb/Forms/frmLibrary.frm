VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6.ocx"
Begin VB.Form frmLibrary 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Library"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   688
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   4515
      Left            =   0
      MousePointer    =   1  'Arrow
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4845
         TabIndex        =   7
         Top             =   0
         Width           =   4230
      End
      Begin MMPlayerXProject.Button BTN 
         Height          =   255
         Index           =   0
         Left            =   465
         TabIndex        =   6
         ToolTipText     =   "Biblioteca"
         Top             =   15
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   450
         ButtonColor     =   255
         Style           =   1
      End
      Begin MSComctlLib.ImageList ImgIconos 
         Left            =   1290
         Top             =   3510
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":21FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":4451
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":6683
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":8867
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":AAA3
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":CCF3
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":EECB
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":111B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":133DB
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":155B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":1782E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":19AA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLibrary.frx":1BCE0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VSFlex6Ctl.vsFlexGrid FGS 
         Bindings        =   "frmLibrary.frx":1DF2A
         DragIcon        =   "frmLibrary.frx":1DF3E
         Height          =   1170
         Left            =   45
         TabIndex        =   4
         Top             =   2565
         Visible         =   0   'False
         Width           =   795
         _cx             =   1402
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
      Begin VB.FileListBox fPlayList 
         Height          =   480
         Left            =   1800
         Pattern         =   "*.pls"
         TabIndex        =   2
         Top             =   3315
         Visible         =   0   'False
         Width           =   585
      End
      Begin MSComctlLib.TreeView TreeFiles 
         Height          =   3780
         Left            =   0
         TabIndex        =   1
         Top             =   315
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   6668
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImgIconos"
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VSFlex6Ctl.vsFlexGrid FG 
         Bindings        =   "frmLibrary.frx":1E248
         DragIcon        =   "frmLibrary.frx":1E25C
         Height          =   3810
         Left            =   2595
         TabIndex        =   3
         Top             =   330
         Width           =   5880
         _cx             =   10372
         _cy             =   6720
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
      Begin MMPlayerXProject.Button BTN 
         Height          =   255
         Index           =   3
         Left            =   9150
         TabIndex        =   8
         ToolTipText     =   "Buscar"
         Top             =   0
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   450
         ButtonColor     =   255
         Style           =   1
      End
      Begin MMPlayerXProject.Button BTN 
         Height          =   255
         Index           =   2
         Left            =   2580
         TabIndex        =   10
         ToolTipText     =   "Eliminar PlayList"
         Top             =   15
         Visible         =   0   'False
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   450
         ButtonColor     =   255
         Style           =   1
      End
      Begin MMPlayerXProject.Button BTN 
         Height          =   255
         Index           =   1
         Left            =   1245
         TabIndex        =   11
         ToolTipText     =   "Actualizar"
         Top             =   30
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   450
         ButtonColor     =   255
         Style           =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   4125
         TabIndex        =   9
         Top             =   30
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   4215
         Width           =   7305
      End
   End
End
Attribute VB_Name = "frmLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnnMusic As ADODB.Connection
Dim CMD As ADODB.Command
Dim iContFilas As Integer
Dim bCtrlSel As Boolean
Dim bShiftSel As Boolean
Dim i As Integer
Dim j As Integer
Dim iSelRow As Integer

Dim cWindows As New cWindowSkin
Dim cAjustarDesk As New clsDockingHandler
Dim InFormDrag As Boolean

Sub Contruir_Arbol()
 Dim sClave As String
 Dim sLastNode As String
 Dim sLAlbum As String
 Dim sLArtist As String
 Dim sNode As String
 Dim sAlbum As String
 Dim sArtist As String
 Dim sGenre As String
 Dim rsFiles As New ADODB.Recordset
 
 On Error GoTo HELL
  TreeFiles.Nodes.Clear
  TreeFiles.Nodes.Add , , "kAll", "Local Audio", 3
  TreeFiles.Nodes.Add , , "kMediaLibrary", "Media Library", 5
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kAlbum", "Album", 6
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kArtist\Album", "Artist\Album", 7
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kCDMedia", "CD Media", 8
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kPath", "File Location", 9
  TreeFiles.Nodes.Add "kPath", tvwChild, "kFullPath", "Full Path", 9
  TreeFiles.Nodes.Add "kPath", tvwChild, "kFolder", "by Folder", 2
  TreeFiles.Nodes.Add "kMediaLibrary", tvwChild, "kGenre", "Genre", 13
  TreeFiles.Nodes.Add , , "kPlayList", "Play List", 14
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kTopHits", "Top Hits", 12
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kRecP", "Recently Played", 11
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kRecI", "Recently Imported", 10
  TreeFiles.Nodes.Add "kPlayList", tvwChild, "kPlaL", "Play Lists", 14
  
  
 CMD.CommandText = "SELECT DISTINCT GENRE,ARTIST FROM MUSIC WHERE ONCD =FALSE ORDER BY GENRE"
 
 Set rsFiles = CMD.Execute
  
 Dim rsArtist As New ADODB.Recordset
 Dim rsAlbum As New ADODB.Recordset

''  // GENRE
    Do Until rsFiles.EOF
       sGenre = CStr(Trim(rsFiles!Genre))
       If sGenre = "" Then sGenre = "Desconocido"
       
       If "GAAGE" & LCase(sGenre) <> sLastNode Then
           sLastNode = "GAAGE" & LCase(sGenre)
           TreeFiles.Nodes.Add "kGenre", tvwChild, sLastNode, sGenre, 13
       End If
       rsFiles.MoveNext
    Loop
  
  
 CMD.CommandText = "SELECT DISTINCT ALBUM FROM MUSIC WHERE ONCD=FALSE ORDER BY ALBUM"
 Set rsFiles = CMD.Execute
 sArtist = ""
 sAlbum = ""
  
''  // ALBUM
    Do Until rsFiles.EOF
       sAlbum = CStr(Trim(rsFiles!Album))
       If sAlbum = "" Then sAlbum = "Desconocido"
       
       If "A  AL" & LCase(sAlbum) <> sLastNode Then
           sLastNode = "A  AL" & LCase(sAlbum)
           TreeFiles.Nodes.Add "kAlbum", tvwChild, sLastNode, sAlbum, 6
       End If
       rsFiles.MoveNext
    Loop

'  // ARTIST - ALBUMS

 
 CMD.CommandText = "SELECT DISTINCT ARTIST FROM MUSIC WHERE ONCD=FALSE ORDER BY ARTIST"
 Set rsFiles = CMD.Execute
   sLastNode = ""
   sLAlbum = ""
     Do Until rsFiles.EOF
       sArtist = CStr(Trim(rsFiles!Artist))
       If sArtist = "" Then sArtist = "Desconocido"
       If "AA AR" & LCase(sArtist) <> sLastNode Then
           sLastNode = "AA AR" & LCase(sArtist)
           TreeFiles.Nodes.Add "kArtist\Album", tvwChild, sLastNode, sArtist, 7
           CMD.CommandText = "SELECT DISTINCT ALBUM FROM MUSIC WHERE ONCD=FALSE AND ARTIST='" & rsFiles!Artist & "'"
           Set rsAlbum = CMD.Execute
           If rsAlbum.RecordCount > 1 Then
           Do Until rsAlbum.EOF
                sAlbum = CStr(Trim(rsAlbum!Album))
                If sAlbum = "" Then sAlbum = "Desconocido"
                If "AA AL|" & LCase(sArtist) & "|" & LCase(sAlbum) <> sLAlbum Then
                sLAlbum = "AA AL|" & LCase(sArtist) & "|" & LCase(sAlbum)
                TreeFiles.Nodes.Add sLastNode, tvwChild, sLAlbum, sAlbum, 6
                End If
                rsAlbum.MoveNext
           
           Loop
           End If
           rsAlbum.Close
       
       End If
       rsFiles.MoveNext
    Loop
    
 '// FILE LOCATION
 Dim sKey As String, s As String
 Dim sPath() As String

 On Error Resume Next

 CMD.CommandText = "SELECT DISTINCT FILEPATH FROM MUSIC WHERE ONCD=FALSE"
 Set rsFiles = CMD.Execute
 sLastNode = ""
 sArtist = ""
 sAlbum = ""

 '// add albums folders
 Do Until rsFiles.EOF
    s = rsFiles!FilePath
        
    sPath = Split(s, "\", , vbTextCompare)
    TreeFiles.Nodes.Add "kFolder", tvwChild, "FL FO" & CStr(s & "\"), sPath(UBound(sPath)), 2
        
    If sLastNode <> sPath(0) Then
       TreeFiles.Nodes.Add "kFullPath", tvwChild, "FL CA" & CStr(sPath(0) & "\"), sPath(0), 1
       sLastNode = sPath(0)
    End If
    
    sKey = "FL CA" & sPath(0) & "\"
    
    For i = 1 To UBound(sPath)
       'If TreeFiles.Nodes(sKey).Children = 0 Then
          TreeFiles.Nodes.Add sKey, tvwChild, sKey & sPath(i) & "\", sPath(i), 2
        'End If
      If i = UBound(sPath) Then
        sKey = sKey & sPath(i)
      Else
        sKey = sKey & sPath(i) & "\"
      End If
    Next i
    rsFiles.MoveNext
    sKey = ""
 Loop
 
 
 '// CD MEDIA
 
 CMD.CommandText = "SELECT DISTINCT FILEPATH FROM MUSIC WHERE ONCD=TRUE"
 Set rsFiles = CMD.Execute
 sLastNode = ""
 sArtist = ""
 sAlbum = ""

 '// add albums folders
 Do Until rsFiles.EOF
    s = rsFiles!FilePath
        
    sPath = Split(s, "\", , vbTextCompare)
    
    If sLastNode <> sPath(0) Then
       TreeFiles.Nodes.Add "kCDMedia", tvwChild, "CDMCA" & CStr(sPath(0) & "\"), sPath(0), 1
       sLastNode = sPath(0)
    End If
    
    sKey = "CDMCA" & sPath(0) & "\"
    
    For i = 1 To UBound(sPath)
       'If TreeFiles.Nodes(sKey).Children = 0 Then
          TreeFiles.Nodes.Add sKey, tvwChild, sKey & sPath(i) & "\", sPath(i), 2
        'End If
      If i = UBound(sPath) Then
        sKey = sKey & sPath(i)
      Else
        sKey = sKey & sPath(i) & "\"
      End If
    Next i
    rsFiles.MoveNext
    sKey = ""
 Loop
 
 '-----------------------------------------------------------------------------------
'// buskar los archivos de playlist y agragarlos
fPlayList.Pattern = "*.pls"


If Dir(tAppConfig.AppConfig & "Library\", vbDirectory) <> "" Then
  fPlayList.Path = tAppConfig.AppConfig & "Library\"
  
  For i = 0 To fPlayList.ListCount - 1
      TreeFiles.Nodes.Add "kPlaL", tvwChild, "kPlaE" & Left(fPlayList.List(i), Len(fPlayList.List(i)) - 4), Left(fPlayList.List(i), Len(fPlayList.List(i)) - 4), 14
  Next i
End If

 
' '// AGREGAR CD ROMS Y OTROS
'    Dim FS As New FileSystemObject
'    Dim dDrive As Drive
'    Dim dDrives As Drives
'
'
'    Set dDrives = FS.Drives
'
'    For Each dDrive In dDrives
'       'If dDrive.IsReady = True Then
'          Select Case dDrive.DriveType
'
'             Case 0 '/* Desconocido
'             Case 1 '/* Separable
'             Case 2 '/* Fijo
''                cboDrives.AddItem dDrive.DriveLetter & ": [" & dDrive.VolumeName & "]"
'             Case 3 '/* Red
'             Case 4 '/* CDROM
'                 If dDrive.IsReady = True Then
'                     sGenre = dDrive.DriveLetter & ": [" & dDrive.VolumeName & "]"
'                 Else
'                    sGenre = dDrive.DriveLetter
'                 End If
'
'                 TreeFiles.Nodes.Add "kCDS", tvwChild, "CDSFI" & dDrive.DriveLetter, sGenre, 1
'
'             Case 5 '/* Disco RAM
'          End Select
'      ' End If
'    Next
'
' Set FS = Nothing
TreeFiles.Nodes("kMediaLibrary").Expanded = True
TreeFiles.Nodes("kPlayList").Expanded = True

rsFiles.Close
rsArtist.Close
rsAlbum.Close
Set rsFiles = Nothing
Set rsArtist = Nothing
Set rsAlbum = Nothing

Exit Sub
HELL:
MsgBox err.Description

End Sub

Sub Guardar_PlayList(GuardarComo As Boolean)
 Dim Fnum As Integer, j As Integer
 Dim archivoINI As String
 Dim intClave As Integer
 Dim s As String
 On Error GoTo BITCH
 
 If Dir(tAppConfig.AppConfig & "Library\", vbDirectory) = "" Then Exit Sub
 
 If GuardarComo = True Then
    s = InputBox(LineLanguage(251), "Play List", "PlayList_" & TreeFiles.Nodes("kPlaL").Children, Me.Left + (Me.Width / 2) - 2000, Me.Top + 2000)
    If Trim(s) = "" Then Exit Sub
    archivoINI = tAppConfig.AppConfig & "Library\" & s & ".pls"
 Else
    archivoINI = tAppConfig.AppConfig & "Library\" & TreeFiles.SelectedItem.Text & ".pls"
 End If

 

If Dir(archivoINI) <> "" Then '// si existe el archivo borrarlo
 SetAttr archivoINI, vbNormal
 Kill archivoINI
End If
    
FG.SaveGrid archivoINI, flexFileData
    
If GuardarComo = True Then
   TreeFiles.Nodes.Add "kPlaL", tvwChild, "kPlaE" & s, s, 14
   frmPopUp.Cargar_Lista_Rep
End If
Exit Sub
BITCH:
MsgBox err.Description
End Sub

Sub Cargar_PlayListTracks(sPlayList As String)
 Dim sFileVis As String
 On Error Resume Next
 sFileVis = tAppConfig.AppConfig & "Library\" & sPlayList & ".pls"
 If Dir(sFileVis) = "" Then Exit Sub
 FG.LoadGrid sFileVis, flexFileData

End Sub

Sub Agregar_Todos(bNuevo As Boolean)
On Error GoTo HELL

Dim iRows As Integer
Dim X As Integer
  
  If FG.Rows = 1 Then Exit Sub
      
  If bNuevo = True Then
     frmPlayList.PL.Clear
     frmPlayList.PL.Rows = 0
  End If
        
  iRows = frmPlayList.PL.Rows
  If iRows < 0 Then iRows = 0
  frmPlayList.PL.Rows = frmPlayList.PL.Rows + (FG.Rows - 1)
  For i = iRows To frmPlayList.PL.Rows - 1
      For j = 0 To FG.Cols - 1
          frmPlayList.PL.TextMatrix(i, j + 2) = FG.TextMatrix(X + 1, j)
      Next j
      X = X + 1
      frmPlayList.Agregar_PlayList_de_BaseDatos i
  Next i
HELL:
End Sub

Sub Eliminar_Biblioteca()
On Error GoTo HELL
Dim X As Integer
Dim sSQL As String
If FG.Rows <= 1 Then Exit Sub

If FGS.Row = -1 Then
    sSQL = Replace(FG.TextMatrix(FG.Row, FG.Cols - 1), "'", "''", , , vbTextCompare)
    CMD.CommandText = "DELETE FROM MUSIC WHERE FILE='" & sSQL & "'"
    FG.RemoveItem (FG.Row)
    CMD.Execute
    Exit Sub
End If
FGS.Col = 0
FGS.Sort = flexSortNumericAscending

For i = 0 To FGS.Rows - 1
    sSQL = Replace(FG.TextMatrix(CInt(FGS.TextMatrix(i, 0)) - X, FG.Cols - 1), "'", "''", , , vbTextCompare)
    CMD.CommandText = "DELETE FROM MUSIC WHERE FILE='" & sSQL & "'"
    FG.RemoveItem CInt(FGS.TextMatrix(i, 0)) - X
    CMD.Execute
    X = X + 1
Next
FG.Select CInt(FGS.TextMatrix(i - 1, 0)), 0
FGS.Rows = 0

HELL:
End Sub

Sub Eliminar_Archivos_Biblioteca()
On Error GoTo HELL
Dim rs As New ADODB.Recordset
rs.Open "SELECT FILE FROM MUSIC", cnnMusic, adOpenDynamic, adLockOptimistic

Do Until rs.EOF
   If Dir(rs!File) = "" Then
      rs.Delete
   End If
   rs.MoveNext
Loop
rs.Update
rs.Close
Set rs = Nothing
HELL:
End Sub


Sub Explorar_Archivos()
On Error Resume Next
Dim strPathExplore As String
If FG.Rows <= 1 Then Exit Sub
strPathExplore = FG.TextMatrix(FG.Row, FG.Cols - 1)
strPathExplore = Left(strPathExplore, InStrRev(strPathExplore, "\"))

Shell "explorer.exe " & strPathExplore, vbMaximizedFocus

End Sub

Sub BTN_Click(Index As Integer)

On Error GoTo HELL
Dim sCampos As String
Dim archivoINI As String
Dim stipo As String
Dim sSQL As String
Select Case Index
    Case 0
        PopupMenu frmPopUp.mnuBiblioteca
    Case 3
        If Len(Trim(Text1)) = 0 Then Exit Sub
        Label1.Caption = ""
        sCampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"
        sSQL = Replace(Text1, "'", "''", , , vbTextCompare)
        CMD.CommandText = "SELECT " & sCampos & " FROM MUSIC WHERE TITLE LIKE '%" & sSQL & "%' OR ARTIST LIKE '%" & sSQL & "%' ORDER BY TITLE"
        Set FG.DataSource = CMD.Execute
        FG.ColHidden(FG.Cols - 1) = True
    Case 2
      '//CLICK EN PLAYLISTS
        stipo = Left(TreeFiles.SelectedItem.Key, 5)
        If stipo <> "kPlaE" Then Exit Sub
         sCampos = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
         If Dir(tAppConfig.AppConfig & "Library\", vbDirectory) = "" Then Exit Sub
         archivoINI = tAppConfig.AppConfig & "Library\" & sCampos & ".pls"
        If Dir(archivoINI) <> "" Then '// si existe el archivo borrarlo
            SetAttr archivoINI, vbNormal
            Kill archivoINI
            TreeFiles.Nodes.Remove TreeFiles.SelectedItem.Key
            FG.Clear
        End If
    Case 1
        Contruir_Arbol

End Select
HELL:
'MsgBox err.Description
End Sub
Sub Agregar_Seleccionadas(bNuevo As Boolean)
On Error Resume Next
Dim iRows As Integer
Dim X As Integer

If FG.Rows = 1 Then Exit Sub
If bNuevo = True Then
    frmPlayList.PL.Clear
    frmPlayList.PL.Rows = 0
End If

If FGS.Row = -1 Then
    FG_DblClick
    Exit Sub
End If

iRows = frmPlayList.PL.Rows
If iRows < 0 Then iRows = 0
frmPlayList.PL.Rows = frmPlayList.PL.Rows + FGS.Rows
For i = iRows To frmPlayList.PL.Rows - 1
    For j = 0 To FG.Cols - 1
        frmPlayList.PL.TextMatrix(i, j + 2) = FG.TextMatrix(CInt(FGS.TextMatrix(X, 0)), j)
    Next j
    X = X + 1
    frmPlayList.Agregar_PlayList_de_BaseDatos i
Next i
End Sub

Sub Editar_Archivos()
On Error Resume Next
If FG.Rows = 1 Then Exit Sub
If FGS.Row = -1 Then
    frmTags.listRef.ListItems.Clear
    frmTags.fileTags.Clear
    frmTags.Load_Tags FG.TextMatrix(CInt(FG.Row), FG.Cols - 1)
    frmTags.Show
    Exit Sub
End If

    frmTags.listRef.ListItems.Clear
    frmTags.fileTags.Clear
For i = 0 To FGS.Rows - 1
    frmTags.Load_Tags FG.TextMatrix(CInt(FGS.TextMatrix(i, 0)), FG.Cols - 1)
Next
    frmTags.Show
End Sub

Private Sub FG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Shift = 2 Then
    bCtrlSel = True
Else
    bCtrlSel = False
End If

End Sub

Private Sub FG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 1
FG.MousePointer = 1

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
    FGS.Text = FG.Row
    iContFilas = iContFilas + 1
        FG.Col = 0
    FG.CellBackColor = FG.BackColorSel
      FG.CellForeColor = FG.ForeColorSel
    FG.Col = 1
    FG.CellBackColor = FG.BackColorSel
      FG.CellForeColor = FG.ForeColorSel
    FG.Col = 2
    FG.CellBackColor = FG.BackColorSel
      FG.CellForeColor = FG.ForeColorSel
    FG.Col = 3
    FG.CellBackColor = FG.BackColorSel
      FG.CellForeColor = FG.ForeColorSel
              FG.Col = 4
    FG.CellBackColor = FG.BackColorSel
    FG.CellForeColor = FG.ForeColorSel
        FG.Col = 5
    FG.CellBackColor = FG.BackColorSel
    FG.CellForeColor = FG.ForeColorSel
        FG.Col = 6
    FG.CellBackColor = FG.BackColorSel
    FG.CellForeColor = FG.ForeColorSel
        FG.Col = 7
    FG.CellBackColor = FG.BackColorSel
    FG.CellForeColor = FG.ForeColorSel
    End If
End If

End Sub

Private Sub Form_Load()
On Error Resume Next
Set cnnMusic = New ADODB.Connection

  With cnnMusic
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Properties("Data Source") = tAppConfig.AppConfig & "Library\music.mdb"
    '.Properties("Jet OLEDB:Database Password") = "Licenciao159"
    .CursorLocation = adUseClient
    .Open
  End With
  
  Set CMD = New ADODB.Command
  CMD.ActiveConnection = cnnMusic

  Contruir_Arbol
  bolMediaLibraryShow = True
  Set cWindows.FormularioPadre = Me
  Set cAjustarDesk.ParentForm = Me
 
End Sub

Sub cargar_formulario()
Dim iX As Integer, iY As Integer

  cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
  iX = Read_INI("Configuration", "ExitButtonX", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\LIBRARY\config.ini")
  iY = Read_INI("Configuration", "ExitButtonY", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\LIBRARY\config.ini")
  cWindows.ButtonExitXY CLng(iX), CLng(iY)
  cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\LIBRARY\config.ini")
  cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\LIBRARY\config.ini")
 
 cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\LIBRARY\"
    picClientArea.Left = cWindows.AreaLeft
    picClientArea.Top = cWindows.AreaTop
    picClientArea.Width = cWindows.AreaWidth
    picClientArea.Height = cWindows.AreaHeight
    
    BTN(0).Left = cWindows.AreaLeft
    BTN(0).Top = 0
    BTN(1).Left = BTN(0).Left + BTN(0).Width
    BTN(1).Top = 0
    BTN(2).Left = BTN(1).Left + BTN(1).Width
    BTN(2).Top = 0
    
    TreeFiles.Top = BTN(0).Height
    TreeFiles.Height = cWindows.AreaHeight - BTN(0).Height - Label1.Height
    
    FG.Top = BTN(0).Height
    FG.Left = TreeFiles.Width + 3
    FG.Height = cWindows.AreaHeight - BTN(0).Height - Label1.Height
    FG.Width = picClientArea.ScaleWidth - TreeFiles.Width - 3
    
    Label1.Top = cWindows.AreaHeight - Label1.Height
    Label1.Width = cWindows.AreaWidth


End Sub

Private Sub Form_Unload(Cancel As Integer)

cnnMusic.Close
Set cnnMusic = Nothing
Set cWindows = Nothing
Set cAjustarDesk = Nothing

End Sub



Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then BTN_Click 3
End Sub

Sub TreeFiles_Click()
Dim color As Single
Dim sAlbum As String
Dim sArtist As String
Dim sGenre As String
Dim sSQL As String
Dim stipo As String
Dim aEle() As String
Dim sCampos As String
Dim sWhere As String
On Error Resume Next

  stipo = Left(TreeFiles.SelectedItem.Key, 5)
    
  If stipo <> "kPlaE" And stipo <> "kRecI" And stipo <> "kRecP" And stipo <> "kTopH" And stipo <> "FL FO" And stipo <> "CDMCA" And stipo <> "kAll" And stipo <> "A  AL" And stipo <> "AA AR" And stipo <> "AA AL" And stipo <> "FL CA" And stipo <> "GAAGE" Then Exit Sub
    
  sCampos = "TITLE,ARTIST,ALBUM,GENRE,YEAR,LENGTH,PLAYCOUNT,PLAYEDLAST,FILE"
'sCampos = "*"
    
  If stipo = "kAll" Then
    sWhere = "WHERE ONCD=FALSE"
    sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK EN ALBUMS
  If stipo = "A  AL" Then
     sAlbum = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
     sWhere = "WHERE ONCD=FALSE AND ALBUM='" & sAlbum & "'"
     sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK EN ARTISTAS
  If stipo = "AA AR" Then
     sArtist = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
     sWhere = "WHERE ONCD=FALSE AND ARTIST='" & sArtist & "'"
     sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK EN ARTISTAS - ALBUM
  If stipo = "AA AL" Then
     aEle = Split(TreeFiles.SelectedItem.Key, "|", , vbTextCompare)
     sArtist = aEle(1)
     sAlbum = aEle(2)
     sWhere = "WHERE ONCD=FALSE AND ARTIST='" & sArtist & "' AND ALBUM='" & sAlbum & "'"
     sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
  End If
  
  '// CLICK EN FILE LOCATION
  If stipo = "FL CA" Or stipo = "FL FO" Then
     If TreeFiles.SelectedItem.Children = 0 Then
        sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
        sGenre = Left(sGenre, Len(sGenre) - 1)
        sWhere = "WHERE ONCD=FALSE AND FILEPATH='" & sGenre & "'"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
     
     End If
  End If
  
  '// CDMEDIA
  If stipo = "CDMCA" Then
     If TreeFiles.SelectedItem.Children = 0 Then
        sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
        sGenre = Left(sGenre, Len(sGenre) - 1)
        sWhere = "WHERE ONCD=TRUE AND FILEPATH='" & sGenre & "'"
        sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
     Else
        If Len(TreeFiles.SelectedItem.Key) = 8 Then
           sGenre = Right(TreeFiles.SelectedItem.Key, 3)
           sWhere = "WHERE ONCD=TRUE AND DRIVE='" & sGenre & "'"
           sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
        End If
     End If
  End If

  
   '//CLICK EN GENEROS
  If stipo = "GAAGE" Then
     sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
     sWhere = "WHERE ONCD=FALSE AND GENRE='" & sGenre & "'"
     sSQL = "SELECT " & sCampos & " FROM MUSIC " & sWhere
  End If
  
  '//CLICK EN TOP HITS
  If stipo = "kTopH" Then
      sWhere = "WHERE ONCD=FALSE ORDER BY PLAYCOUNT DESC"
      sSQL = "SELECT TOP 20 PLAYCOUNT," & sCampos & " FROM MUSIC " & sWhere
  End If
    
    
  '//CLICK EN RECIENTE REPRODUCIDAS
  If stipo = "kRecP" Then
      sWhere = "WHERE ONCD=FALSE ORDER BY PLAYEDLAST DESC"
      sSQL = "SELECT TOP 20 PLAYEDLAST," & sCampos & " FROM MUSIC " & sWhere
  End If
    
  '//CLICK EN IMPORTADAS RECIENTEMENTE
  If stipo = "kRecI" Then
      sWhere = "WHERE ONCD=FALSE ORDER BY LASTUPDATE DESC"
      sSQL = "SELECT TOP 20 LASTUPDATE," & sCampos & " FROM MUSIC " & sWhere
  End If
  BTN(2).Visible = False
  '//CLICK EN PLAYLISTS
  If stipo = "kPlaE" Then
      sGenre = Right(TreeFiles.SelectedItem.Key, Len(TreeFiles.SelectedItem.Key) - 5)
      Cargar_PlayListTracks sGenre
      BTN(2).Visible = True
      Exit Sub
  End If

    iSelRow = 0
      
    
    CMD.CommandText = sSQL
    Set FG.DataSource = CMD.Execute
    FG.ColHidden(FG.Cols - 1) = True
    
    Label1.Caption = "RECORDS:[ " & FG.Rows - 1 & " ]   -   "
    
    
    '// ESTADISTICAS
    Dim rs As New ADODB.Recordset
    Dim KILOBYTES As Long, BYTES As Long
    Dim DD As Long, HH As Long, MM As Long, SS As Long, sTempTime As String, lSeconds As Long
    
    rs.Open "SELECT SUM(BYTES) AS TOTAL, SUM(SECONDS) AS TIEMPO  FROM MUSIC " & sWhere, cnnMusic, adOpenForwardOnly, adLockReadOnly
    BYTES = rs!total
    
    KILOBYTES = CLng(BYTES / 1024 * 100) / 100
    sTempTime = "SIZE: [ " & Format(KILOBYTES, "000,000") & " KB. ]   -   "
    
    Label1.Caption = Label1.Caption & sTempTime
    sTempTime = ""
    lSeconds = rs!TIEMPO
    DD = lSeconds \ 86400     ' Dias
    lSeconds = Abs(lSeconds - (DD * 86400))
    HH = lSeconds \ 3600      ' Horas
    MM = lSeconds \ 60 Mod 60 ' Minutos
    SS = lSeconds Mod 60      ' Segundos
    sTempTime = "TIME:[ "
    If DD > 0 Then sTempTime = sTempTime & DD & " dias. "
    If HH > 0 Then sTempTime = sTempTime & HH & " Hr. "
    Label1.Caption = Label1.Caption & sTempTime & MM & " Min. " & Format$(SS, "00") & " Sec. ]"
    rs.Close

End Sub

Sub Actualizar_Track(sFile As String)
Dim rsAct As New ADODB.Recordset
Dim iContar As Integer
On Error GoTo HELL
Dim s As String

s = Replace(sFile, "'", "''", , , vbTextCompare)
rsAct.Open "SELECT PLAYCOUNT,PLAYEDLAST FROM MUSIC WHERE FILE='" & s & "'", cnnMusic, adOpenDynamic, adLockPessimistic

If rsAct.RecordCount = 1 Then
    'rsAct!PlayCount = rsAct!PlayCount + 1
    'rsAct!PLayedLast = Now()
    'rsAct.UpdateBatch adAffectCurrent
    iContar = rsAct!PlayCount + 1
    CMD.CommandText = "UPDATE MUSIC SET PLAYCOUNT=" & iContar & ",PLAYEDLAST='" & Now() & "' WHERE FILE='" & s & "'"
    CMD.Execute
End If
    rsAct.Close
    Set rsAct = Nothing
HELL:
'MsgBox err.Description
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
    
    BTN(0).Left = cWindows.AreaLeft
    BTN(0).Top = 0
    BTN(1).Left = BTN(0).Left + BTN(0).Width
    BTN(1).Top = 0
    BTN(2).Left = BTN(1).Left + BTN(1).Width
    BTN(2).Top = 0
    
    TreeFiles.Top = BTN(0).Height
    TreeFiles.Height = cWindows.AreaHeight - BTN(0).Height - Label1.Height
    
    FG.Top = BTN(0).Height
    FG.Left = TreeFiles.Width + 3
    FG.Height = cWindows.AreaHeight - BTN(0).Height - Label1.Height
    FG.Width = picClientArea.ScaleWidth - TreeFiles.Width - 3
    
    Label1.Top = cWindows.AreaHeight - Label1.Height
    Label1.Width = cWindows.AreaWidth
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
frmMain.Mostrar_Media_Library
End If
InFormDrag = False
End Sub


Private Sub picClientArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Me.MousePointer = vbDefault
 
 
 
 If Button = vbLeftButton And picClientArea.MousePointer = 9 Then
  If X > 100 And X < picClientArea.Width - 100 Then
   TreeFiles.Width = X - 2
   FG.Left = X + 2
   FG.Width = picClientArea.ScaleWidth - TreeFiles.Width - 2
  End If
End If
 If X > TreeFiles.Left + TreeFiles.Width And X < TreeFiles.Left + TreeFiles.Width + 6 Then
   picClientArea.MousePointer = 9
 Else
   picClientArea.MousePointer = vbDefault
 End If

End Sub

Private Sub Form_Paint()
cWindows.Formulario_Paint
End Sub

Private Sub FG_DblClick()
   Dim iTrack As Integer
   Dim bEureka As Boolean
   If FG.Row = 0 Then Exit Sub
    
    FG.Col = FG.Cols - 1
    
    If Dir(FG.Text) = "" Then Exit Sub
    
    For iTrack = 0 To frmPlayList.PL.Rows - 1
      If LCase(FG.Text) = LCase(frmPlayList.PL.TextMatrix(iTrack, 10)) Then
        bEureka = True
        Exit For
      End If
    Next
    
    If bEureka = False Then
      frmPlayList.PL.Rows = frmPlayList.PL.Rows + 1
      For j = 0 To FG.Cols - 1
           frmPlayList.PL.TextMatrix(frmPlayList.PL.Rows - 1, j + 2) = FG.TextMatrix(FG.Row, j)
      Next j
      frmPlayList.Agregar_PlayList_de_BaseDatos frmPlayList.PL.Rows - 1

      frmPlayList.Play_Track frmPlayList.PL.Rows - 1
    Else
      frmPlayList.Play_Track iTrack
    End If
  
End Sub
Private Sub FG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo HELL
Dim color As Single
Dim iFila As Integer

If Button = vbRightButton Then
   PopupMenu frmPopUp.mnuMenubiblioteca
   Exit Sub
End If

If bCtrlSel = False And FGS.Rows = 0 Then
 Dim NRows As Integer, iStart As Integer, iEnd As Integer, iTemp As Integer
  NRows = Abs(FG.RowSel - FG.Row)
  If NRows = 0 Then Exit Sub
  iStart = FG.Row
  iEnd = FG.RowSel
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
       If FG.Text = FGS.Text Then Exit Sub
     Next i
    
    FGS.Col = 0
    FGS.Rows = iContFilas + 1
    FGS.Row = iContFilas
    FGS.Text = FG.Row
    iContFilas = iContFilas + 1
    
    '// Seleccionar fila
    
    FG.Col = 0
    FG.CellBackColor = FG.BackColorSel
      FG.CellForeColor = FG.ForeColorSel
    FG.Col = 1
    FG.CellBackColor = FG.BackColorSel
      FG.CellForeColor = FG.ForeColorSel
    FG.Col = 2
    FG.CellBackColor = FG.BackColorSel
      FG.CellForeColor = FG.ForeColorSel
    FG.Col = 3
    FG.CellBackColor = FG.BackColorSel
    FG.CellForeColor = FG.ForeColorSel
        FG.Col = 4
    FG.CellBackColor = FG.BackColorSel
    FG.CellForeColor = FG.ForeColorSel
        FG.Col = 5
    FG.CellBackColor = FG.BackColorSel
    FG.CellForeColor = FG.ForeColorSel
        FG.Col = 6
    FG.CellBackColor = FG.BackColorSel
    FG.CellForeColor = FG.ForeColorSel
        FG.Col = 7
    FG.CellBackColor = FG.BackColorSel
    FG.CellForeColor = FG.ForeColorSel
    iFila = FG.Row
    Exit Sub
End If




If bCtrlSel = False And FGS.Rows > 0 And bShiftSel = False Then
  iFila = FG.Row
   For i = 0 To FGS.Rows - 1
      FGS.Row = i
      
      If CInt(FGS.Text) Mod 2 <> 0 Then
        color = FG.BackColor
      Else
        color = FG.BackColorAlternate
      End If
      FG.Row = CInt(FGS.Text)
      FG.Col = 0
      FG.CellBackColor = color
      FG.CellForeColor = FG.ForeColor
      FG.Col = 1
      FG.CellBackColor = color
      FG.CellForeColor = FG.ForeColor
      FG.Col = 2
      FG.CellBackColor = color
      FG.CellForeColor = FG.ForeColor
      FG.Col = 3
      FG.CellBackColor = color
      FG.CellForeColor = FG.ForeColor
      FG.Col = 4
      FG.CellBackColor = color
      FG.CellForeColor = FG.ForeColor
      FG.Col = 5
      FG.CellBackColor = color
      FG.CellForeColor = FG.ForeColor
      FG.Col = 6
      FG.CellBackColor = color
      FG.CellForeColor = FG.ForeColor
      FG.Col = 7
      FG.CellBackColor = color
      FG.CellForeColor = FG.ForeColor
    Next i
    FGS.Rows = 0
    FG.Row = iFila
    bCtrlSel = False
    Exit Sub
End If

bShiftSel = False
FGS.Rows = 0
Exit Sub
HELL:

End Sub



Private Sub TreeFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = vbDefault
 
End Sub
