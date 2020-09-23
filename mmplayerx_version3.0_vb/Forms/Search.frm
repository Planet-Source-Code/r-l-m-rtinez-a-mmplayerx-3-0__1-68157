VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search and add tracks"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Search.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstRef 
      Height          =   840
      Left            =   210
      TabIndex        =   10
      Top             =   3435
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.DirListBox DirSearch 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   450
      TabIndex        =   9
      Top             =   2460
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.FileListBox FileSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00400000&
      Height          =   420
      Hidden          =   -1  'True
      Left            =   1425
      MousePointer    =   99  'Custom
      Pattern         =   "*.mp3;*.wma"
      System          =   -1  'True
      TabIndex        =   8
      Top             =   2775
      Visible         =   0   'False
      Width           =   1350
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   330
      Left            =   30
      TabIndex        =   6
      Top             =   1365
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Start Now"
      Height          =   330
      Left            =   1290
      TabIndex        =   3
      Top             =   1710
      Width           =   2970
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   330
      Left            =   3945
      TabIndex        =   1
      Top             =   75
      Width           =   1605
   End
   Begin VB.ComboBox cboDrives 
      Height          =   315
      ItemData        =   "Search.frx":000C
      Left            =   1065
      List            =   "Search.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   75
      Width           =   2850
   End
   Begin VB.Label lblTracks 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15
      TabIndex        =   7
      Top             =   1110
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label lblProgress 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   15
      TabIndex        =   5
      Top             =   495
      Width           =   1995
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Look in:"
      Height          =   195
      Left            =   315
      TabIndex        =   4
      Top             =   120
      Width           =   690
   End
   Begin VB.Label lblProgress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   885
      Index           =   0
      Left            =   15
      TabIndex        =   2
      Top             =   750
      Width           =   5535
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As ADODB.Recordset
Dim cnn As ADODB.Connection
Dim strSQL As String
Dim cFile As New cMP3
Dim iMp3Totales As Integer


Dim i As Integer, j As Integer
Dim bCancel As Boolean
Dim sLastPath As String

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'|  BUSKEDA METODO UNO: MAS RAPIDO PERO UTILIZANDO OBJETOS DIR Y FILE :)                 |
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Search_Files(strPath As String)
 On Error GoTo HELL
 Dim strPathCur As String
 Dim bEncontro As Boolean
 Dim XXX As Integer
 '// Primero buscar en el directorio padre para buscar despues en subdirectorios
 If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
 
 
 '// set pather at files list box
 If strPathern = "" Then strPathern = "*.mp3"
   
  FileSearch.Pattern = strPathern
  
  FileSearch.Path = strPath
 If FileSearch.ListCount > 0 Then
       iMp3Totales = iMp3Totales + FileSearch.ListCount
       lstRef.AddItem strPath
 End If
 
 '// poner cursor de busqueda si hay del skin
 strPathCur = tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\"
 If Dir(strPathCur & "curFind.cur") <> "" Then
   frmMain.picNormalMode.MouseIcon = LoadPicture(strPathCur & "curFind.cur")
 End If
  
    '// Empezar ha buskar
    Call Start_Search(strPath)
    
  
Exit Sub
HELL:
 If Dir(strPathCur & "curMain.cur") <> "" Then frmMain.picNormalMode.MouseIcon = LoadPicture(strPathCur & "curMain.cur")
End Sub


'// metod for search is very faster

Sub Start_Search(strPath As String)
 On Error Resume Next  '// manejador de error por si permisos de acceso a los directorios
 
 DoEvents '// para que deje trabajar el Windows
 Dim subdirs As Integer, k As Integer, intFolder As Integer
 ReDim subdirs_name(0 To 10) As String  '// arreglo para directorios
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
        lstRef.AddItem FileSearch.Path
        iMp3Totales = iMp3Totales + FileSearch.ListCount
        lblProgress(1).Caption = "Files: [ " & iMp3Totales & " ]"
      End If
Next intFolder

'//-----------Buscamos en subdirectorios ----------------------------------------
'// como es una procedimento que se llama a si mismo las variables anteriores
'// se siguen conservando hasta que termine
For k = 0 To subdirs - 1
 '// mostramos los directorios de busqueda
 
 lblProgress(0).Caption = subdirs_name(k)
 Start_Search subdirs_name(k)
Next

End Sub


Private Sub Add_Files()
  Dim iProg As Integer
  Dim FS As New FileSystemObject
  Dim bCDROM As Boolean
  Dim dDrive As drive
  Dim dDrives As Drives
  Dim lSeconds As Long
  Dim sFile As String, sTitle As String, sArtist As String, sAlbum As String, sGenre As String, sYear As String, sComment As String
  
    
  On Error Resume Next
  
  lblProgress(0).Visible = False
  lblTracks.Visible = True

  If bCancel = True Then GoTo HELL
  
  pbProgress.min = 0
  pbProgress.Max = iMp3Totales
  pbProgress.Value = 0
  lblTracks.Caption = "adding [ " & iMp3Totales & " ] tracks "
  lblProgress(1).Caption = ""
  pbProgress.Visible = True
   
   cFile.Read_MPEGInfo = True
    For i = 0 To lstRef.ListCount - 1
       DoEvents
       
       If bCancel = True Then GoTo HELL

        FileSearch.Path = lstRef.List(i)
       '// chekar si esta en un cdrom u otro
        Set dDrive = FS.Drives(Left(FileSearch.Path, 1))
        If dDrive.DriveType = CDRom Then
           bCDROM = True
        Else
            bCDROM = False
        End If
        
        For j = 0 To FileSearch.ListCount - 1
        
            cFile.Read_File_Tags FileSearch.Path & "\" & FileSearch.List(j)
            
            sTitle = Replace(cFile.Title, "'", " ", , , vbTextCompare)
            sArtist = Replace(cFile.Artist, "'", " ", , , vbTextCompare)
            sAlbum = Replace(cFile.Album, "'", " ", , , vbTextCompare)
            sYear = Replace(cFile.Year, "'", " ", , , vbTextCompare)
            sGenre = Replace(cFile.Genre, "'", " ", , , vbTextCompare)
            sComment = Replace(cFile.Comment, "'", " ", , , vbTextCompare)
            
            If sTitle = "" Then sTitle = Left(FileSearch.List(j), Len(FileSearch.List(j)) - 4)
            If sArtist = "" Then sArtist = "Unknow"
            If sAlbum = "" Then sAlbum = "Unknow"
            If sYear = "" Then sYear = Year(Now())
            If sGenre = "" Then sGenre = "Other"
            If sComment = "" Then sComment = "Uncomment"
            
            rst.AddNew
            rst!File = FileSearch.Path & "\" & FileSearch.List(j)
            rst!Title = sTitle
            rst!Artist = sArtist
            rst!Album = sAlbum
            rst!Year = sYear
            rst!Genre = sGenre
            rst!Comments = sComment
            rst!length = cFile.MPEG_DurationTime
            rst!BYTES = cFile.FileSize
            rst!Seconds = cFile.DurationInSecs
'            rst!LastUpdate = cFile.LastUpdate
            rst!Quality = cFile.Quality
            rst!Situation = cFile.Situation
            rst!Mood = cFile.Mood
            rst!FilePath = FileSearch.Path
            rst!OnCD = bCDROM
            rst!drive = Left(FileSearch.Path, 3)
            rst.Update
            iProg = iProg + 1
            pbProgress.Value = iProg
        Next j
    Next i
    
  lblTracks.Caption = "Ready: [ " & iMp3Totales & " ] tracks found"
  pbProgress.Visible = False
  cmdSearch.Caption = "Start Now"

 Exit Sub
HELL:
  cmdSearch.Caption = "Start Now"
  lblTracks.Caption = "Cancel by user"
  pbProgress.Visible = False
End Sub


Private Sub cboDrives_Click()
   lblProgress(0).Caption = "All folders"
   sLastPath = "All folders"
End Sub

Private Sub cmdBrowse_Click()
 On Error GoTo HELL
 Dim sPath As String
  sPath = Explorador_Para_Directorios(Me.hwnd, LineLanguage(234))

  If sPath = "" Then Exit Sub
  For i = 2 To cboDrives.ListCount - 1
    If LCase(Left(sPath, 1)) = LCase(Left(cboDrives.List(i), 1)) Then
      cboDrives.ListIndex = i
      Exit For
    End If
  Next i
  lblProgress(0).Caption = sPath
  sLastPath = sPath
Exit Sub
HELL:

End Sub



Private Sub cmdSearch_Click()
    Dim FS As New FileSystemObject
    Dim dDrive As drive
    Dim dDrives As Drives
  On Error Resume Next
   
  If cboDrives.ListIndex < 0 Then Exit Sub
   
  If bSearching = True Then
     bSearching = False
     GoTo BITCH
  End If
   
  
  cmdSearch.Caption = LineLanguage(233)
  cboDrives.Enabled = False
  cmdBrowse.Enabled = False
  lblTracks.Visible = False
  lblProgress(0).Visible = True
  bCancel = False
  lstRef.Clear
  iMp3Totales = 0

  
  '/* Search in All Hard Drives
  If cboDrives.ListIndex = 0 Then
        
     Set dDrives = FS.Drives
     bSearching = True
     For Each dDrive In dDrives
       If dDrive.IsReady = True Then
          If dDrive.DriveType = Fixed Then
            Search_Files dDrive.Path & "\"
          End If
       End If
    Next
    Call Add_Files
    bSearching = False
    GoTo BITCH
  End If
  
  '/* Search in All Drives
  If cboDrives.ListIndex = 1 Then
        
     Set dDrives = FS.Drives
     bSearching = True
     For Each dDrive In dDrives
       If dDrive.IsReady = True Then
          If dDrive.DriveType = Fixed Or dDrive.DriveType = CDRom Then
             Search_Files dDrive.Path & "\"
          End If
       End If
    Next
    Call Add_Files
    bSearching = False
    GoTo BITCH
  End If
  
  '/* search in other hard disk
  If cboDrives.ListIndex > 1 And lblProgress(0).Caption = "All folders" Then
    bSearching = True
    Search_Files Left(cboDrives.List(cboDrives.ListIndex), 1) & ":\"
    Call Add_Files
    bSearching = False
    GoTo BITCH
  End If
  
  '/* search in folder
  If lblProgress(0).Caption <> "All folders" Then
    If Dir(lblProgress(0).Caption, vbDirectory) <> "" Then
       bSearching = True
       Search_Files lblProgress(0).Caption
       Call Add_Files
       bSearching = False
    End If
  End If
  
  frmLibrary.Contruir_Arbol
BITCH:
     lblProgress(0).Caption = sLastPath
     cboDrives.Enabled = True
     cmdBrowse.Enabled = True
     cmdSearch.Caption = LineLanguage(232)

End Sub


Private Sub Form_Load()
  On Error Resume Next
    Dim FS As New FileSystemObject
    Dim dDrive As drive
    Dim dDrives As Drives
    
     Set rst = New ADODB.Recordset
    Set cnn = New ADODB.Connection
    With cnn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Data Source") = tAppConfig.AppConfig & "Library\music.mdb"
        '.Properties("Jet OLEDB:Database Password") = "Licenciao159"
        .CursorLocation = adUseClient
        .Open
    End With
  
    strSQL = "SELECT * FROM Music"
    
    rst.Open strSQL, cnn, adOpenDynamic, adLockOptimistic


    Set dDrives = FS.Drives
      
    cboDrives.AddItem "Local hard drives"
    cboDrives.AddItem "All Drives"
    
    For Each dDrive In dDrives
       If dDrive.IsReady = True Then
          Select Case dDrive.DriveType
              
             Case 0 '/* Desconocido
             Case 1 '/* Separable
             Case 2 '/* Fijo
                cboDrives.AddItem dDrive.DriveLetter & ": [" & dDrive.VolumeName & "]"
             Case 3 '/* Red
             Case 4 '/* CDROM
               cboDrives.AddItem dDrive.DriveLetter & ": [" & dDrive.VolumeName & "]"
             Case 5 '/* Disco RAM
          End Select
       End If
    Next
  Load_Language_Search
  bolSearchShow = True
  Me.Icon = frmMain.Icon
  Me.Left = (Screen.Width - Me.Width) / 2 '// centrar form
  Me.Top = (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
  If bSearching = True Then
     If bSearching = True Then bSearching = False
  End If
  bolSearchShow = False
  
  cnn.Close
  Set cnn = Nothing
End Sub
