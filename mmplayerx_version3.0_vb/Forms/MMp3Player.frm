VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "MMPlayerX v. 3.0"
   ClientHeight    =   4170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5730
   ForeColor       =   &H00000000&
   Icon            =   "MMp3Player.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3480
      Top             =   3495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MMp3Player.frx":1AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MMp3Player.frx":1E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MMp3Player.frx":21A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MMp3Player.frx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MMp3Player.frx":284A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMiniMode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00642909&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   -15
      MousePointer    =   99  'Custom
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   274
      TabIndex        =   10
      Top             =   2820
      Visible         =   0   'False
      Width           =   4110
      Begin MMPlayerXProject.ScrollText ScrollText 
         Height          =   90
         Index           =   4
         Left            =   195
         TabIndex        =   13
         Top             =   45
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   159
         PictureText     =   "MMp3Player.frx":2B9E
         CaptionText     =   "00:00"
         ScrollVelocity  =   150
         AutoSize        =   -1  'True
      End
      Begin MMPlayerXProject.ScrollText ScrollText 
         Height          =   90
         Index           =   5
         Left            =   645
         TabIndex        =   17
         Top             =   45
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   159
         CaptionText     =   "track name"
         AlignText       =   2
         ScrollType      =   1
         ScrollVelocity  =   150
         Scroll          =   -1  'True
      End
      Begin MMPlayerXProject.Button ButtonMini 
         Height          =   150
         Index           =   0
         Left            =   2580
         TabIndex        =   33
         Top             =   15
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button ButtonMini 
         Height          =   150
         Index           =   1
         Left            =   2745
         TabIndex        =   34
         Top             =   15
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button ButtonMini 
         Height          =   150
         Index           =   2
         Left            =   2910
         TabIndex        =   35
         Top             =   15
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button ButtonMini 
         Height          =   150
         Index           =   3
         Left            =   3075
         TabIndex        =   36
         Top             =   15
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button ButtonMini 
         Height          =   150
         Index           =   4
         Left            =   3240
         TabIndex        =   37
         Top             =   15
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button ButtonMini 
         Height          =   150
         Index           =   5
         Left            =   15
         TabIndex        =   38
         Top             =   15
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button ButtonMini 
         Height          =   150
         Index           =   6
         Left            =   3585
         TabIndex        =   39
         Top             =   15
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button ButtonMini 
         Height          =   150
         Index           =   7
         Left            =   3750
         TabIndex        =   40
         Top             =   15
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button ButtonMini 
         Height          =   150
         Index           =   8
         Left            =   3915
         TabIndex        =   41
         Top             =   15
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Slider Slider 
         Height          =   90
         Index           =   2
         Left            =   600
         TabIndex        =   42
         Top             =   195
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   159
         BackColor       =   65535
         Position        =   1
      End
      Begin MMPlayerXProject.Slider Slider 
         Height          =   90
         Index           =   3
         Left            =   2205
         TabIndex        =   43
         Top             =   195
         Width           =   870
         _ExtentX        =   159
         _ExtentY        =   1535
         BackColor       =   65535
         Max             =   255
         Position        =   1
      End
   End
   Begin VB.TextBox txtSTIcon 
      Height          =   285
      Index           =   1
      Left            =   2130
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   3765
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtSTIcon 
      Height          =   285
      Index           =   2
      Left            =   1230
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   3765
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtSTIcon 
      Height          =   285
      Index           =   0
      Left            =   945
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   3765
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtSTIcon 
      Height          =   285
      Index           =   3
      Left            =   1545
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   3765
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtSTIcon 
      Height          =   285
      Index           =   4
      Left            =   1830
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   3765
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.ListBox LyricsRef 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2490
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox picWallOriginal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2775
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   3
      Top             =   3765
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picWallProp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3135
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   2
      Top             =   3795
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2445
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3780
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picNormalMode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   186
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   343
      TabIndex        =   0
      Top             =   0
      Width           =   5145
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1785
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   51
         TabIndex        =   44
         Top             =   2415
         Visible         =   0   'False
         Width           =   765
      End
      Begin MMPlayerXProject.Slider Slider 
         Height          =   150
         Index           =   0
         Left            =   15
         TabIndex        =   31
         Top             =   1800
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   265
         BackColor       =   65535
         Position        =   1
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   270
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   2055
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   476
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.ScrollText ScrollText 
         Height          =   90
         Index           =   0
         Left            =   75
         TabIndex        =   12
         Top             =   1335
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   159
         CaptionText     =   "00:00"
         ScrollVelocity  =   150
         AutoSize        =   -1  'True
      End
      Begin VB.PictureBox picSpectrum 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00008000&
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
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   825
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   11
         Top             =   1185
         Width           =   1335
      End
      Begin MMPlayerXProject.ScrollText ScrollText 
         Height          =   90
         Index           =   2
         Left            =   570
         TabIndex        =   14
         Top             =   1200
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   159
         PictureText     =   "MMp3Player.frx":4CDA
         CaptionText     =   "128"
         ScrollVelocity  =   150
         AutoSize        =   -1  'True
      End
      Begin MMPlayerXProject.ScrollText ScrollText 
         Height          =   90
         Index           =   3
         Left            =   585
         TabIndex        =   15
         Top             =   1350
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   159
         PictureText     =   "MMp3Player.frx":6E16
         CaptionText     =   "44"
         ScrollVelocity  =   150
         AutoSize        =   -1  'True
      End
      Begin MMPlayerXProject.ScrollText ScrollText 
         Height          =   90
         Index           =   1
         Left            =   30
         TabIndex        =   16
         Top             =   1650
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   159
         CaptionText     =   "track name"
         AlignText       =   2
         ScrollType      =   1
         ScrollVelocity  =   150
         Scroll          =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   270
         Index           =   1
         Left            =   555
         TabIndex        =   19
         Top             =   2055
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   476
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   270
         Index           =   2
         Left            =   930
         TabIndex        =   20
         Top             =   2055
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   476
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   270
         Index           =   3
         Left            =   1305
         TabIndex        =   21
         Top             =   2055
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   476
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   270
         Index           =   4
         Left            =   1680
         TabIndex        =   22
         Top             =   2055
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   476
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   195
         Index           =   5
         Left            =   315
         TabIndex        =   23
         Top             =   930
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   344
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   195
         Index           =   6
         Left            =   735
         TabIndex        =   24
         Top             =   930
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   344
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   195
         Index           =   7
         Left            =   1155
         TabIndex        =   25
         Top             =   930
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   344
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   195
         Index           =   8
         Left            =   1575
         TabIndex        =   26
         Top             =   930
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   344
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   150
         Index           =   12
         Left            =   30
         TabIndex        =   27
         Top             =   30
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   150
         Index           =   13
         Left            =   4575
         TabIndex        =   28
         Top             =   30
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   150
         Index           =   14
         Left            =   4740
         TabIndex        =   29
         Top             =   30
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   150
         Index           =   15
         Left            =   4905
         TabIndex        =   30
         Top             =   30
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Slider Slider 
         Height          =   1815
         Index           =   1
         Left            =   2235
         TabIndex        =   32
         Top             =   435
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   3201
         BackColor       =   65535
         Max             =   255
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   150
         Index           =   10
         Left            =   795
         TabIndex        =   45
         ToolTipText     =   "Lista Reproduccion"
         Top             =   495
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   150
         Index           =   9
         Left            =   270
         TabIndex        =   46
         ToolTipText     =   "visualizacion Studio"
         Top             =   510
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin MMPlayerXProject.Button Button 
         Height          =   150
         Index           =   11
         Left            =   1410
         TabIndex        =   47
         ToolTipText     =   "Biblioteca multimedia"
         Top             =   525
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   265
         ButtonColor     =   65535
         MaskColor       =   16711935
         MousePointer    =   99
         Style           =   1
         UseMaskColor    =   -1  'True
      End
      Begin VB.Image ImgCaratula 
         Appearance      =   0  'Flat
         Height          =   2085
         Left            =   2430
         Stretch         =   -1  'True
         Top             =   225
         Width           =   2610
      End
   End
   Begin VB.Timer Timer_Intro 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   150
      Top             =   3240
   End
   Begin VB.Timer Timer_Player 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   615
      Top             =   3255
   End
   Begin VB.Timer Timer_Crossfade 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1035
      Top             =   3240
   End
   Begin VB.Timer Timer_Wait 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   1485
      Top             =   3240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===========================================================================
'   Proyect : Music Mp3 Player X
'   Version : 3.0
'   Author  : Raúl Martínez
'   Email   : escorpio36@hotmail.com
'   Web     : www.geocities.com/skoria_36
'   Update  : January 2007, Valle de Santiago, Guanajuato, México
'
'   You do NOT have rights to redistribute this code, in whole or in part
'   without my permission.  You also may not recompile the code and release
'   it as another program without my permission.  If you would like to modify
'   this code and distribute it in either as source code or as a compiled
'   program please contact me at [escorpio36@hotmail.com] before doing so.
'   I would appreciate being notified of any modifications even if you do not
'   intend to redistribute it... (OR YOU WILL BURN IN THE HELL)
'
'   Components:
'     - FMOD.dll  version 3.73 (in app path)
'     - Microsoft Common Dialog Control 6.0
'     - Microsoft Windows Common Control 5.0
'     - Microsoft Windows Common Control 6.0
'     - :-) VideoSoft VSFlex 6.0 Controls (OLEDB)
'
'   References:
'     - Microsoft Scripting Runtime.
'     - Microsoft Activex Data Objects 2.0 Library
'
'   Any idea, comment, suggestions, doubts, bugs, skins, languages, etc.
'   please email me.
'
'P.D.
'  ----------------------------------------------------
'  * Si NoS PinTaN CoMo UnOs GuEvOnEs, No Lo SoMoS ...*
'  *       ¡¡¡ ViVa MeXiCo KabRonEs !!!               *
'  *      QuE Se SiEnTa El PoWeR MeXiCaNo...          *
'  ----------------------------------------------------
'
'=============================================================================


'=============================================================================
' WARNING: THIS PROGRAM USE SUBCLASSING, SO...
'          DO NOT PRESS THE STOP BUTTON IN VISUAL BASIC IDE!!!!
'=============================================================================


Dim ttDemo                As New Tooltip
Dim cAjustarDesk As New clsDockingHandler
Dim InFormDrag As Boolean
Dim sArtFile As String

Public sSysTrayText       As String
Dim PlayerIntro           As Boolean
Dim TiempoIntro           As Integer
Dim PlayerLoop            As Boolean
Dim PlayerMute            As Boolean
'-----------------------------------------------
Public PlayerIsPlaying    As String   '// Estado del Player
Public VolumeNActuaL      As Long     '// Volumen del Reproductor
'------------------------------------------------
Dim bolAleatorio     As Boolean  '// Orden Aleatorio
Dim AleatorioRola()       As Integer  '// arreglo para Aleatorio

Dim bSlider               As Boolean  '// arrastrando slider posbar

'// Variables para mostrar el Karaoke
Public LyricsIndex        As Integer

'// Variables para la minimascara
Dim bolTimeAct            As Boolean

'// Spectrum
Dim arryPeaks(50)         As Single
Dim arryWaitPeak(50)      As String

'// Crossfade funcion
Dim lCurrentChannel            As Long
Dim lVol                  As Long
Dim lChannelOut           As Long
Dim lChannelIn             As Long

'##################################################
' VARIABLES PARA ANKLAR A LA LISTA DE REPRODUCCION
'/* Variables para la minimascara
Dim bolDragMini As Boolean
Dim StartDragX As Single, StartDragY As Single
Dim mAttachedToRight As Boolean
Dim mAttachedToLeft As Boolean
Dim mAttachedToTop As Boolean
Dim mAttachedToBottom As Boolean
Dim mSnapDistance As Long


Dim iPosHook As Integer


Sub Show_Mensaje(Mensaje As String)
  If bMiniMask = True Then
     ScrollText(5).CaptionText = Mensaje
  Else
     ScrollText(1).CaptionText = Mensaje
  End If
  Timer_Wait.Enabled = True
  
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Search_Caratula(sPath As String)
 On Error Resume Next
 Dim miNombre As String, sPathFront As String
 Dim bolEureka As Boolean, bolCaratula As Boolean
 Dim sPathOther As String
 Dim strPath As String

If sArtFile <> "" Then
  If Dir(sArtFile) <> "" Then bolCaratula = True
   sPathFront = sArtFile
   GoTo ENDSEARCH
End If

strPath = Left(sPath, InStrRev(sPath, "\"))

'// buskar archivos .JPG
 miNombre = Dir(strPath & "*.jpg")
 Do While miNombre <> ""
      sPathOther = strPath & miNombre
      bolEureka = LCase(Trim(miNombre)) Like "*caratula*"
      If bolEureka = False Then bolEureka = LCase(Trim(miNombre)) Like "*portada*"
      If bolEureka = False Then bolEureka = LCase(Trim(miNombre)) Like "*front*"
      If bolEureka = False Then bolEureka = LCase(Trim(miNombre)) Like "*frt*"
      If bolEureka = True Then
        bolCaratula = True
        sPathFront = strPath & miNombre
        GoTo ENDSEARCH
      End If
    miNombre = Dir
 Loop

'// buskar archivos .BMP
 miNombre = Dir(strPath & "*.bmp")
 Do While miNombre <> ""
      sPathOther = strPath & miNombre
      bolEureka = LCase(Trim(miNombre)) Like "*caratula*"
      If bolEureka = False Then bolEureka = LCase(Trim(miNombre)) Like "*portada*"
      If bolEureka = False Then bolEureka = LCase(Trim(miNombre)) Like "*front*"
      If bolEureka = False Then bolEureka = LCase(Trim(miNombre)) Like "*frt*"
      If bolEureka = True Then
        bolCaratula = True
        sPathFront = strPath & miNombre
        GoTo ENDSEARCH
      End If
    miNombre = Dir
 Loop

ENDSEARCH:

If Trim(sPathOther) <> "" And bolCaratula = False Then
  bolCaratula = True
  sPathFront = sPathOther
End If

'// si enkuentra alguna caratula
If bolCaratula = True Then
  If LCase(sPathFront) = LCase(strRutaCaratula) Then Exit Sub
    ImgCaratula.Stretch = True
    ImgCaratula.Picture = LoadPicture(sPathFront)
    strRutaCaratula = sPathFront
  
'    If bolCaratulaShow = True Then ' si esta cargado el frmcaratula mostrar la caratula
      frmCaratula.Picture1.Picture = LoadPicture(sPathFront)
      frmCaratula.Mover_Form
'    End If
    
    If bolVisShow = True Then frmSpectrum.Setup_Visualizacion
       
    If frmOpciones.chkWallpaper.Value = 1 Then ConfigurarWallpaper
    
Else
'    If bolCaratulaShow = True Then 'si esta caragado y no tiene caratula mostrar la default
      frmCaratula.Picture1.Picture = frmPopUp.picDefaultLogo.Picture
      frmCaratula.Mover_Form
'    End If
    
    strRutaCaratula = ""
    
    If bolVisShow = True Then frmSpectrum.Setup_Visualizacion
    
    If frmOpciones.chkWallpaper.Value = 1 Then ConfigurarWallpaper
    ImgCaratula.Picture = LoadPicture("")
End If

End Sub

Sub Play_Crossfade()
 On Error Resume Next
 Dim lngX As Long

    If (lCurrentChannel = 0) Then
        lCurrentChannel = 1:  lChannelIn = 1:  lChannelOut = 0
    Else
        lCurrentChannel = 0:  lChannelIn = 0:  lChannelOut = 1
    End If
        
        Stream_Open sFileMainPlaying, FSOUND_NORMAL, lCurrentChannel, True, VolumeNActuaL
        'Stream_SetVolume lCurrentChannel, 0
        If PlayerMute = True Then Stream_SetMute lCurrentChannel, True
        
        lVol = VolumeNActuaL
        peCrossFadeType = CrossfadeNormal
        Timer_Crossfade.Interval = iCrossfadeTrack
        Timer_Crossfade.Enabled = True
        
End Sub
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Play()
 On Error Resume Next
 If Dir(sFileMainPlaying) = "" Or frmPlayList.PL.Rows = 0 Then Exit Sub
 If sFileMainPlaying = "" Then
    frmPlayList.Play_Track 0
 End If
  If PlayerIntro = True Then Timer_Intro.Enabled = True: TiempoIntro = 0
  If PlayerIsPlaying = "pause" Then Pause_Play: Exit Sub
  'If PlayerIsPlaying = "true" Then Five_Seg_Forward: Exit Sub
  
  '// check if player int frmtags is playing
  If PlayerState = "true" Or PlayerState = "pause" Then frmTags.Stop_Player
  
  Timer_Player.Enabled = False

  Start_Play
  Search_Caratula sFileMainPlaying
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Start_Play()
On Error GoTo error
Dim lngVolume As Long

   Load_File_Tags
   
   Start_Lyrics
   
   
           
   If iCrossfadeTrack <> 0 Then
     Play_Crossfade
   Else

     Stream_Open sFileMainPlaying, FSOUND_NORMAL, lCurrentChannel, True, VolumeNActuaL
     If PlayerMute = True Then Stream_SetMute lCurrentChannel, True
     'Stream_SetVolume lCurrentChannel, VolumeNActuaL
     
   End If
   If PlayerIsPlaying <> "true" Then frmOpciones.chkDSP_Click 1
   Timer_Player.Enabled = True
   
   PlayerIsPlaying = "true"
   Image_State_Rep
    
     Slider(2).Max = CInt(Stream_GetDuration(lCurrentChannel))
     Slider(0).Max = CInt(Stream_GetDuration(lCurrentChannel))
   
Exit Sub
error:
   Stop_Player
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Stop_Player()
 If PlayerIsPlaying = "false" Then Exit Sub
  
  If bolVisShow = False Then frmMain.Stop_Draw_Spectrum
  
  If bolVisShow = True Then frmSpectrum.Stop_Visualizacion
 
 On Error Resume Next
  
  PlayerIsPlaying = "false"
  Timer_Player.Enabled = False
  Image_State_Rep
    
  If PlayerIntro = True Then Timer_Intro.Enabled = False
  'POR SI TIENE ALGUN EFECTO ACTIVADO Y NO SE HAGA LENTO
  
  If iCrossfadeStop <> 0 Then
    'fade out
    lVol = VolumeNActuaL
    peCrossFadeType = FadeIn
    Timer_Crossfade.Interval = iCrossfadeStop
    Timer_Crossfade.Enabled = True
  Else
    Stream_Stop lCurrentChannel
    FX_Disable
  End If
  
  If bMiniMask = True Then
     ScrollText(4).CaptionText = "00:00"
     Slider(2).Value = 0
  Else
     ScrollText(0).CaptionText = "00:00"
     Slider(0).Value = 0
  End If
  
   
  If tCurrentID3.Has_Lyrics3_Tag = True And bolLyricsShow = True Then frmLyrics.Reset_Values
 
  
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Pause_Play()
 Dim CurState As Long
 
' If ListRep.ListCount = 0 Or TotalAlbumS = 0 Then Exit Sub
 
  If PlayerIsPlaying = "false" Then Exit Sub
     CurState = Stream_GetState(lCurrentChannel)
     
     '// Esta Reproduciendo, pausar
     If CurState = 2 Then
       If PlayerIntro = True Then Timer_Intro.Enabled = False
       PlayerIsPlaying = "pause"
       Image_State_Rep
       
       If iCrossfadeStop <> 0 Then
         '- Fade in -------------------------------------------------------
         lVol = VolumeNActuaL
         peCrossFadeType = FadeIn
         Timer_Crossfade.Interval = iCrossfadeStop
         Timer_Crossfade.Enabled = True
         '-----------------------------------------------------------------
       Else
         Stream_Pause lCurrentChannel
         FX_Disable
       End If
        
     Else
     '// Si esta pausado, reproducir
       PlayerIsPlaying = "true"
       Stream_Pause lCurrentChannel
       If iCrossfadeStop <> 0 Then
         Stream_SetVolume lCurrentChannel, 0
         '- Fade Out -------------------------------------------------------
         lVol = 0
         peCrossFadeType = FadeOut
         Timer_Crossfade.Interval = iCrossfadeStop
         Timer_Crossfade.Enabled = True
         '-----------------------------------------------------------------
       End If
       frmOpciones.chkDSP_Click 0
       Image_State_Rep
       If PlayerIntro = True Then Timer_Intro.Enabled = True
     End If
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Next_Track()
  If frmPlayList.PL.Rows = 0 Then Exit Sub
  
  If frmPopUp.mnuOrdenAleatorio.Checked = True Then
    Randomize_Order
    Exit Sub
  End If
    
  frmPlayList.Next_Track
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Previous_Track()
  If frmPlayList.PL.Rows = 0 Then Exit Sub
  
  If frmPopUp.mnuOrdenAleatorio.Checked = True Then
    Randomize_Order
    Exit Sub
  End If
  
  frmPlayList.Previous_Track
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Five_Seg_Forward()
 On Error Resume Next
 Dim CurPos As Long
'  If ListRep.ListCount = 0 Or PlayerIsPlaying <> "true" Then Exit Sub
  
  CurPos = Stream_GetPosition(lCurrentChannel)
  CurPos = CurPos + 5
  If CurPos > Stream_GetDuration(lCurrentChannel) Then CurPos = Stream_GetDuration(lCurrentChannel)
  Stream_SetPosition lCurrentChannel, CurPos
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Five_Seg_Backward()
 On Error Resume Next
 Dim CurPos As Long
'  If ListRep.ListCount = 0 Or PlayerIsPlaying <> "true" Then Exit Sub
  CurPos = Stream_GetPosition(lCurrentChannel)
  CurPos = CurPos - 5
  If CurPos < 0 Then CurPos = 0
  Stream_SetPosition lCurrentChannel, CurPos
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Intro()
  If PlayerIntro = False Then
    'poner intro activado
    Button(5).Selected = True
    '-------------------------------------------
    PlayerIntro = True
    TiempoIntro = 0
    Timer_Intro.Enabled = True
    frmPopUp.mnuIntro.Checked = True
    Show_Mensaje "Intro ON"
  Else
   'poner intro desactivado
    Button(5).Selected = False
   '----------------------------------------------
    PlayerIntro = False
    Timer_Intro.Enabled = False
    frmPopUp.mnuIntro.Checked = False
    Show_Mensaje "Intro OFF"
  End If
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Player_Mute()
 On Error Resume Next
  If PlayerMute = False Then
    '--activar silencio --------------------
    Button(6).Selected = True
    PlayerMute = True
    Stream_SetMute lCurrentChannel, True
    frmPopUp.mnuSilencio.Checked = True
    Show_Mensaje "Mute ON"
  Else
    'Desactivar el mute --------------------------------
    Button(6).Selected = False
    PlayerMute = False
    Stream_SetMute lCurrentChannel, False
    frmPopUp.mnuSilencio.Checked = False
    Show_Mensaje "Mute OFF"
  End If
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Player_Repeat()
 If PlayerLoop = False Then
   '---Activar loop -----------------------------
    Button(7).Selected = True
    PlayerLoop = True
    frmPopUp.mnuRepetir.Checked = True
    Show_Mensaje "Repeat ON"
  Else
   '--- Descativar el loop ---------------------------
    Button(7).Selected = False
    PlayerLoop = False
    frmPopUp.mnuRepetir.Checked = False
    Show_Mensaje "Repeat OFF"
  End If
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Randomize_Order()
 On Error Resume Next
  Dim i As Integer, j As Integer
  Dim iMP3 As Integer
  Static stcRolaAleat As Integer

'------- ALEATORIO DE ALBUMS -----------------------------------------------------------
  '// si es la perimera vez
 iMP3 = frmPlayList.PL.Rows
 If bolAleatorio = False Then
     '// redimencionar arreglo con el numero de elementos de la lista de reprod.
     
     ReDim AleatorioRola(iMP3 - 1)
     
     Randomize
          
     If PlayerIsPlaying = "false" Then
       AleatorioRola(0) = Int(iMP3 * Rnd)
     Else
       AleatorioRola(0) = frmPlayList.PL.Row
        If AleatorioRola(0) = -1 Then AleatorioRola(0) = Int(iMP3 * Rnd)
     End If
     
   '// numero de aleatorios a kalkular
   For j = 1 To iMP3 - 1
     DoEvents
      '// skar numero aleatorio
      Randomize
      AleatorioRola(j) = Int(iMP3 * Rnd)
         '// compararlo con los aleatorios anteriores
         '// deskontando el anterior
         For i = 0 To j - 1
            If AleatorioRola(j) = AleatorioRola(i) Then
              j = j - 1
               If j < 1 Then j = 1
              Exit For
            End If
         Next i
    Next j
     bolAleatorio = True
     '// variable para apuntar al numero de arreglo
     stcRolaAleat = 0
     If PlayerIsPlaying = "false" Then
      frmPlayList.Play_Track AleatorioRola(stcRolaAleat)

     End If
     
  '// si no es la primera vez
  Else
    stcRolaAleat = stcRolaAleat + 1
    If stcRolaAleat < iMP3 Then
      frmPlayList.Play_Track AleatorioRola(stcRolaAleat)
    Else
'       Stop_Player
'       Randomize_Click
       stcRolaAleat = 0
       frmPlayList.Play_Track AleatorioRola(stcRolaAleat)
    End If
  End If

End Sub



'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Image_State_Rep()

 If bMiniMask = False Then '// esta en su forma normal
  Button(1).Selected = False
  Button(2).Selected = False
  Button(3).Selected = False
  Select Case PlayerIsPlaying
   Case "true"  'Reproduciendo
     Button(1).Selected = True
   Case "false" 'detenido
     Button(3).Selected = True
   Case "pause" 'Pausado
     Button(2).Selected = True
 End Select
'//------------------------------------------------------------------------
'//------------------------------------------------------------------------
Else '// esta en la minimascara
  ButtonMini(1).Selected = False
  ButtonMini(2).Selected = False
  ButtonMini(3).Selected = False
  Select Case PlayerIsPlaying
   Case "true"  'Reproduciendo
     ButtonMini(1).Selected = True
   Case "false" 'detenido
     ButtonMini(3).Selected = True
   Case "pause" 'Pausado
     ButtonMini(2).Selected = True
 End Select
End If
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Function Convert_Time(ByVal LSec As Long) As String
 Dim HH As Long, MM As Long, SS As Long
 Dim tmp As String
 
 HH = LSec \ 3600  '// calkular horas
 MM = LSec \ 60 Mod 60 '// Calkular minutos
 SS = LSec Mod 60  '// calkular segundos
 
 If HH > 0 Then tmp = HH & ":"
 Convert_Time = tmp & MM & ":" & Format$(SS, "00")
End Function


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Randomize_Click()
 
 If bSearching = True Then Exit Sub
 
  '// desactivar randomize
  If frmPopUp.mnuOrdenAleatorio.Checked = True Then
    bolAleatorio = False
    Button(8).Selected = False
    frmPopUp.mnuOrdenAleatorio.Checked = False
    Show_Mensaje "Random OFF"
  Else
    If frmPlayList.PL.Rows <= 1 Then Exit Sub
      Button(8).Selected = True
      Randomize_Order
      frmPopUp.mnuOrdenAleatorio.Checked = True
      Show_Mensaje "Random ON"
     
  End If
     
End Sub




Public Sub Draw_Spectrum()
  Dim X1 As Single, Y1 As Single
  Dim X2 As Single, Y2 As Single
  Dim iPeak As Single
  Dim i As Integer, iSleep
  Dim Max&
    
 
  
  '// para ahorrar recursos al estar buskando
  If bSearching = True Then Exit Sub
  
  
With picSpectrum
  If bolVisShow = True Then
     Dim sngSpectrumDataVis(200) As Single
     Spectrum_GetData lCurrentChannel, sngSpectrumDataVis()
     frmSpectrum.Update_Visualizacion sngSpectrumDataVis()
     Exit Sub
  End If
  On Error Resume Next
  
  picSpectrum.Cls
  
  If frmPopUp.mnuSpecNone.Checked = True Then Exit Sub

  Dim sngSpectrumData(50) As Single
  Spectrum_GetData lCurrentChannel, sngSpectrumData()
  

   
 If frmPopUp.mnuSpecBars.Checked = True Then ' dibujar el tipo barras
       
        For i = 0 To tSpectrum.iBars
           X1 = i * (.ScaleWidth / tSpectrum.iBars)
           X2 = X1 + ((.ScaleWidth / tSpectrum.iBars) - tSpectrum.iSpacio)
           Y1 = .ScaleHeight
           
           Max = Format(sngSpectrumData(i), ".00") * picSpectrum.ScaleHeight
                       
           '// restringir hasta alto del pic - el alto del peak
           If Max >= Y1 And tSpectrum.bDrawPeaks = True Then Max = Y1 - tSpectrum.iPeakHeight
      
           Y2 = .ScaleHeight - Max
           
           '====================================================================
           '// bars
           If tSpectrum.bDrawBars = True Then
              picSpectrum.Line (X1, Y1)-(X2, Y2), tSpectrum.lBackColorBar, BF
              picSpectrum.Line (X1, Y1)-(X2, Y2), tSpectrum.lLineColorBar, B
           End If
           
           '====================================================================
           '// Peaks
           If tSpectrum.bDrawPeaks = True Then
              '// si cambia la posicion del peak ajustar al alto de la barra
              If arryPeaks(i) < Max Then
                 arryPeaks(i) = Max
                 arryWaitPeak(i) = Time
              End If
              
              '// peak esta abajo
              If arryPeaks(i) < 0 Then arryPeaks(i) = 0
                                 
              iPeak = Y1 - arryPeaks(i)
              
              '// Peak esta al limite superior
              If iPeak <= tSpectrum.iPeakHeight Then iPeak = tSpectrum.iPeakHeight
              
              picSpectrum.Line (X1, iPeak - 1)-(X2, iPeak - tSpectrum.iPeakHeight), tSpectrum.lBackColorPeak, BF
               
              If arryWaitPeak(i) <> "" Then
                '// verificar si todavia sigue alli Peak
                iSleep = DateDiff("s", arryWaitPeak(i), Time)
              End If
              If (iSleep >= 1) Then arryPeaks(i) = arryPeaks(i) - tSpectrum.iPeakGravity
           End If
         
        Next i
        
   ElseIf frmPopUp.mnuSpecOsc.Checked = True Then ' scope
         .CurrentY = .ScaleHeight / 2
         .CurrentX = 0
       
         For i = 0 To tSpectrum.iLinesScope
            X1 = i * (.ScaleWidth / tSpectrum.iLinesScope)
            X2 = X1 + (.ScaleWidth / tSpectrum.iLinesScope)
            Y1 = .ScaleHeight / 2
            Y2 = (sngSpectrumData(i) * Y1)

            picSpectrum.Line Step(0, 0)-(X1 + ((X2 - X1) / 3), Y1 - Y2), tSpectrum.lBackColorScope
            picSpectrum.Line Step(0, 0)-(X1 + (((X2 - X1) / 3) * 2), Y1 + Y2), tSpectrum.lBackColorScope
            picSpectrum.Line Step(0, 0)-(X2, Y1), tSpectrum.lBackColorScope
         Next i

       End If
 End With
 
End Sub

Public Sub Stop_Draw_Spectrum()
  Dim X1 As Single, Y1 As Single
  Dim X2 As Single, Y2 As Single
  Dim iPeak As Single
  Dim i As Integer, iSleep
  Dim Max&
    
 
 picSpectrum.Cls
  
With picSpectrum
  On Error Resume Next
  
  If frmPopUp.mnuSpecNone.Checked = True Then Exit Sub

   
 If frmPopUp.mnuSpecBars.Checked = True Then ' dibujar el tipo barras
       
        For i = 0 To tSpectrum.iBars
           X1 = i * (.ScaleWidth / tSpectrum.iBars)
           X2 = X1 + ((.ScaleWidth / tSpectrum.iBars) - tSpectrum.iSpacio)
           Y1 = .ScaleHeight
           
           Max = 0
                       
           '// restringir hasta alto del pic - el alto del peak
           If Max >= Y1 And tSpectrum.bDrawPeaks = True Then Max = Y1 - tSpectrum.iPeakHeight
      
           Y2 = .ScaleHeight - Max
           
           '====================================================================
           '// bars
           If tSpectrum.bDrawBars = True Then
              picSpectrum.Line (X1, Y1)-(X2, Y2), tSpectrum.lBackColorBar, BF
              picSpectrum.Line (X1, Y1)-(X2, Y2), tSpectrum.lLineColorBar, B
           End If
           
           '====================================================================
           '// Peaks
           If tSpectrum.bDrawPeaks = True Then
              '// peak esta abajo
              arryPeaks(i) = 0
                                 
              iPeak = Y1 - arryPeaks(i)
              
              '// Peak esta al limite superior
              If iPeak <= tSpectrum.iPeakHeight Then iPeak = tSpectrum.iPeakHeight
              
              picSpectrum.Line (X1, iPeak - 1)-(X2, iPeak - tSpectrum.iPeakHeight), tSpectrum.lBackColorPeak, BF
                 
           End If
         
        Next i
        
   ElseIf frmPopUp.mnuSpecOsc.Checked = True Then ' scope
         .CurrentY = .ScaleHeight / 2
         .CurrentX = 0
       
         For i = 0 To tSpectrum.iLinesScope
            X1 = i * (.ScaleWidth / tSpectrum.iLinesScope)
            X2 = X1 + (.ScaleWidth / tSpectrum.iLinesScope)
            Y1 = .ScaleHeight / 2
            Y2 = 0

            picSpectrum.Line Step(0, 0)-(X1 + ((X2 - X1) / 3), Y1 - Y2), tSpectrum.lBackColorScope
            picSpectrum.Line Step(0, 0)-(X1 + (((X2 - X1) / 3) * 2), Y1 + Y2), tSpectrum.lBackColorScope
            picSpectrum.Line Step(0, 0)-(X2, Y1), tSpectrum.lBackColorScope
         Next i

       End If
 End With
 
End Sub





'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub imgCaratula_DblClick()
 bolCaratulaShow = Not bolCaratulaShow
 frmCaratula.Visible = bolCaratulaShow
 frmPopUp.mnuMCaratula.Checked = bolCaratulaShow
End Sub



'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Minimize_Me()
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
   
   If OpcionesMusic.SysTray = False And OpcionesMusic.TaskBar = False Then
      frmPopUp.Visible = True
      OpcionesMusic.TaskBar = True
   End If
   
   If OpcionesMusic.TaskBar = True Then frmPopUp.WindowState = vbMinimized
   If OpcionesMusic.SysTray = True Then Me.Hide
End Sub

Private Sub imgCaratula_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If InFormDrag Then
        Form_Mover X, Y
    End If
End Sub

Private Sub imgCaratula_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
InFormDrag = False
Form_Arriba
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub imgCaratula_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err
 If Button = vbLeftButton Then
     Form_Abajo X, Y
     InFormDrag = True
 End If

 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
Exit Sub
err:
InFormDrag = False
End Sub

Private Sub picNormalMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If InFormDrag Then
        Form_Mover X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub picNormalMode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
InFormDrag = False
Form_Arriba
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub picNormalMode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err
 If Button = vbLeftButton Then
     cAjustarDesk.StartDockDrag X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
     Form_Abajo X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
     InFormDrag = True
 End If

 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
Exit Sub
err:
InFormDrag = False
End Sub

Private Sub picMiniMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If InFormDrag Then
       Form_Mover X * Screen.TwipsPerPixelX, _
            Y * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub picMiniMode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
InFormDrag = False
Form_Arriba
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub picMiniMode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err
 If Button = vbLeftButton Then
     cAjustarDesk.StartDockDrag X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
     Form_Abajo X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
     InFormDrag = True
 End If

 If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
Exit Sub
err:
InFormDrag = False
End Sub




Private Sub picSpectrum_DblClick()
 With frmPopUp
   If .mnuSpecNone.Checked = True Then
      .mnuSpecBars_Click
   ElseIf .mnuSpecBars.Checked = True Then
          .mnuSpecOsc_Click
       ElseIf .mnuSpecOsc.Checked = True Then
              .mnuSpecNone_Click
           End If
        
   End With
End Sub

Private Sub picSpectrum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbRightButton Then PopupMenu frmPopUp.mnuMainSpec
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub ScrollText_DblClick(Index As Integer)
  '// show diferent curent time
  If Index = 0 Or Index = 4 Then bolTimeAct = Not bolTimeAct
  '// stop scroll
  If Index = 1 Or Index = 5 Then ScrollText(Index).StopScroll ScrollText(Index).ScrollingNow
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Button_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 If Button = vbRightButton Then Exit Sub
 
 Select Case Index
   Case 0: Previous_Track
   Case 1: Play
   Case 2: Pause_Play
   Case 3: Stop_Player
   Case 4: Next_Track
   Case 5: Intro
   Case 6: Player_Mute
   Case 7: Player_Repeat
   Case 8: Randomize_Click
   Case 9: Mostrar_Visualizacion
   Case 10: Mostrar_Play_List
   Case 11: Mostrar_Media_Library
   Case 12: Me.PopupMenu frmPopUp.mnuMenuPrincipal
   Case 13: Minimize_Me
   Case 14: Change_Mask True, True: frmMain.Image_State_Rep
   Case 15: Unload Me
 End Select
End Sub

Sub Mostrar_Visualizacion()
If bolVisShow = False Then frmMain.Stop_Draw_Spectrum
bolVisShow = Not bolVisShow
frmPopUp.mnuMaxSpec.Checked = bolVisShow
Button(9).Selected = bolVisShow
frmSpectrum.Visible = bolVisShow
If bolVisShow = True Then
    frmSpectrum.Setup_Visualizacion
    frmSpectrum.Stop_Visualizacion
End If
End Sub



Sub Mostrar_Play_List()
bolPlayListShow = Not bolPlayListShow
frmPopUp.mnuListaR.Checked = bolPlayListShow
Button(10).Selected = bolPlayListShow
frmPlayList.Visible = bolPlayListShow
End Sub

Sub Mostrar_Media_Library()
bolMediaLibraryShow = Not bolMediaLibraryShow
frmPopUp.mnuBibliotecaMultimedia.Checked = bolMediaLibraryShow
Button(11).Selected = bolMediaLibraryShow
frmLibrary.Visible = bolMediaLibraryShow
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Sub Form_KeyPress(KeyAscii As Integer)
 On Error Resume Next
 Dim iValue As Integer
 Dim Index As Integer
 
 Index = 1
 If bMiniMask = True Then Index = 4
 
 If KeyAscii = 45 Then ' - volumen
      VolumeNActuaL = VolumeNActuaL - 4
      If VolumeNActuaL < 0 Then VolumeNActuaL = 0
      If VolumeNActuaL > 255 Then VolumeNActuaL = 255
      Slider(Index).Value = VolumeNActuaL
      Slider_Change Index, VolumeNActuaL
 End If
 If KeyAscii = 43 Then ' + volumen
      VolumeNActuaL = VolumeNActuaL + 4
      If VolumeNActuaL < 0 Then VolumeNActuaL = 0
      If VolumeNActuaL > 255 Then VolumeNActuaL = 255
      Slider(Index).Value = VolumeNActuaL
      Slider_Change Index, VolumeNActuaL
 End If

 
 
 If KeyAscii = 65 Or KeyAscii = 97 Then Five_Seg_Backward 'A Atras 5 seg
 If KeyAscii = 68 Or KeyAscii = 100 Then Five_Seg_Forward 'D Adelante 5 seg
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   
  If KeyCode = 189 Then Timer_Wait.Enabled = True ' - volumen
  If KeyCode = 187 Then Timer_Wait.Enabled = True ' + volumen
  If KeyCode = 90 Then Previous_Track 'Z
  If KeyCode = 88 Then Play 'X
  If KeyCode = 67 Then Pause_Play 'C
  If KeyCode = 86 Then Stop_Player 'V
  If KeyCode = 66 Then Next_Track 'B
  If KeyCode = 73 Then Intro 'I Intro 10 seg
  If KeyCode = 82 Then Player_Repeat 'R Repetir
  If KeyCode = 83 Then Player_Mute 'S Silencio
  If KeyCode = 87 Then Randomize_Click 'W Orden aleatorio coleccion
  If KeyCode = 77 Then frmCaratula.Show  'M Mostrar caratula
  If KeyCode = 70 Then frmSearch.Show  'F Nueva busqueda
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Form_Load()
 On Error Resume Next
  LyricsIndex = 1
  PlayerIsPlaying = "false"
  
  
  '/* note: for debugger is easier comment the next lines (MouseWheel)
  '/*  Call Hook
  '/*  Call UnHook
  
  '/* inizializar detectar la rueda de la rata :)
'  Call Hook
  mSnapDistance = 10 * Screen.TwipsPerPixelX

  Set cAjustarDesk.ParentForm = Me
  Set tCurrentID3 = New cMP3
 On Error GoTo HELL
  FMOD_Initialize 300, 44100, 5, FSOUND_INIT_ENABLESYSTEMCHANNELFX, FSOUND_OUTPUT_DSOUND, FSOUND_MIXER_QUALITY_AUTODETECT, 0
  Spectrum_Enable True

Exit Sub
HELL:
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Form_Unload(Cancel As Integer)
 On Error Resume Next
   
   Stop_Player

   Save_Settings_INI True
   
   If App.PrevInstance = False Then
     If frmOpciones.chkWallpaper.Value = 1 Then PoneRWallPapeROriginaL
   End If
     
     'Borrar el archivo de wallpaper creado si se hizo
   If Dir(DirectoriOWindowS & "MusicMp3.bmp") <> "" Then
     Kill DirectoriOWindowS & "MusicMp3.bmp"
   End If
   
   '/* eliminar monitorizar mause
'   Call Unhook
   Set frmMain = Nothing
   Set ttDemo = Nothing
   Set tCurrentID3 = Nothing
   FMOD_Terminate
   End
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub ButtonMini_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
  If Button = vbRightButton Then Exit Sub
  Select Case Index
    Case 0: Previous_Track
    Case 1: Play
    Case 2: Pause_Play
    Case 3: Stop_Player
    Case 4: Next_Track
    Case 5: Me.PopupMenu frmPopUp.mnuMenuPrincipal
    Case 6: Minimize_Me
    Case 7: Change_Mask False, True: frmMain.Image_State_Rep
    Case 8: Unload Me
  End Select
End Sub

 '+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Slider_Change(Index As Integer, Value As Long)
 Dim intPorcentaje As Integer
 
 On Error Resume Next
 Select Case Index
    Case 1, 3 '// volume bar
               
         intPorcentaje = (Value * 100) / 255

         frmPopUp.mnuVolumen.Caption = LineLanguage(15) & " [ " & intPorcentaje & " % ]"
         
         If bMiniMask = False Then
           ScrollText(1).CaptionText = "Volume " & intPorcentaje & " %"
         Else
           ScrollText(5).CaptionText = "Volume " & intPorcentaje & " %"
         End If
         VolumeNActuaL = Value
         Stream_SetVolume lCurrentChannel, VolumeNActuaL
         
    Case 0, 2 '// pos Bar
         If PlayerIsPlaying = "false" Then Exit Sub
         bSlider = True
         If PlayerIsPlaying = "pause" Then Pause_Play
         
  End Select

End Sub

Private Sub Slider_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' pos bar sliders
If PlayerIsPlaying = "false" Then Exit Sub

If Index = 0 Or Index = 2 Then bSlider = True
End Sub

Private Sub Slider_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 If PlayerIsPlaying = "false" Then Exit Sub
 Select Case Index
    Case 0, 2 '// pos bar
     '// Si esta la minimaskara
       If bMiniMask = True Then
        If bSlider = True Then
         If bolTimeAct = False Then
           ScrollText(4).CaptionText = Convert_Time(Slider(Index).Value)
         Else
           ScrollText(4).CaptionText = "-" & Convert_Time(Slider(Index).Max - Slider(Index).Value)
         End If
        End If
       Else
        If bSlider = True Then
         If bolTimeAct = False Then
           ScrollText(0).CaptionText = Convert_Time(Slider(Index).Value)
         Else
           ScrollText(0).CaptionText = "-" & Convert_Time(Slider(Index).Max - Slider(Index).Value)
         End If
        End If
       End If
     
 End Select

End Sub

Private Sub Slider_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Index
    Case 1, 3 '// volume bar
       Timer_Wait.Enabled = True
    Case 0, 2 '// pos bar
        If PlayerIsPlaying = "false" Then
            If bMiniMask = True Then
               Slider(0).Value = 0
            Else
               Slider(2).Value = 0
            End If
           Exit Sub
        End If
        bSlider = False
        Stream_SetPosition lCurrentChannel, CLng(Slider(Index).Value)
  End Select
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Static rec As Boolean, msg As Long
   
   If bLoading = True Then Exit Sub
   
   msg = X / Screen.TwipsPerPixelX
      ' Captura cada evento de botones del Raton
      Select Case msg
        Case WM_LBUTTONDBLCLK  ' Doble click Boton Izquierdo
         If OpcionesMusic.TaskBar = True Then frmPopUp.WindowState = vbNormal
         
         If OpcionesMusic.SysTray = True Then
           If bolAcercaShow = True Then frmAcerca.Show
           If bolCaratulaShow = True Then frmCaratula.Show
           If bolOpcionesShow = True Then frmOpciones.Show
           If bolLyricsShow = True Then frmLyrics.Show
           If bolTagsShow = True Then frmTags.Show
           If bolVisShow = True Then frmSpectrum.Show
           If bolSearchShow = True Then frmSearch.Show
                      
           Me.Show
           If bolSplashScreen = True Then frmSplash.Show
         End If
           
       Case WM_LBUTTONDOWN  ' Boton Izquierdo pulsado
        Case WM_LBUTTONUP   ' Boton Izquierdo Soltado
        Case WM_RBUTTONDBLCLK ' Doble Click Boton Derecho
        Case WM_RBUTTONDOWN ' Boton derecho pulsado
        Case WM_RBUTTONUP  ' Boton derecho Arriba
           Me.PopupMenu frmPopUp.mnuMenuPrincipal
     End Select
   DoEvents
End Sub




Private Sub Timer_Crossfade_Timer()
 On Error Resume Next
 
 Select Case peCrossFadeType
   Case 0 '// Crossfade normal
         lVol = lVol - 5
         If lVol <= 0 Then
          Stream_Stop lChannelOut
          Timer_Crossfade.Enabled = False
         End If
         
         Stream_SetVolume lChannelOut, lVol
         Stream_SetVolume lChannelIn, Abs(VolumeNActuaL - lVol)
              
   Case 1 '// Fade in
         lVol = lVol - 5
         If lVol <= 0 Then
           FX_Disable
           If PlayerIsPlaying = "false" Then Stream_Stop lCurrentChannel
           If PlayerIsPlaying = "pause" Then Stream_Pause lCurrentChannel
           Timer_Crossfade.Enabled = False
         End If
         
         Stream_SetVolume lCurrentChannel, lVol
         
            
   Case 2 '// Fade Out
         lVol = lVol + 5
         If lVol >= VolumeNActuaL Then Timer_Crossfade.Enabled = False
         
         Stream_SetVolume lCurrentChannel, lVol
         
 End Select
End Sub



'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub txtSTIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Static rec As Boolean, msg As Long
   
   If bLoading = True Then Exit Sub
   
   msg = X / Screen.TwipsPerPixelX
   If rec = False Then
      rec = True
      ' Captura cada evento de botones del Raton
      Select Case msg
        Case WM_LBUTTONUP   ' Boton Izquierdo Soltado
           Select Case Index
               Case 0 '// Previous
                 Previous_Track
               Case 1 '// Play
                 Play
               Case 2 '// Pause
                 Pause_Play
               Case 3 '// Stop
                 Stop_Player
               Case 4 '// Next
                 Next_Track
           End Select
          End Select
      rec = False
   End If
   DoEvents

End Sub

Private Sub Timer_Wait_Timer()
  If bMiniMask = True Then
     ScrollText(5).CaptionText = sTextScroll
  Else
     ScrollText(1).CaptionText = sTextScroll
  End If
  Timer_Wait.Enabled = False
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Timer_Intro_Timer()
 TiempoIntro = TiempoIntro + 1
 If TiempoIntro = 10 Then
  If PlayerLoop = True Then
    Play
  Else
    Next_Track
  End If
  TiempoIntro = 0
 End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Timer_Player_Timer()
  Dim iTimeCross As Integer
  
 '//si esta reproduciendo
  
 If PlayerIsPlaying <> "false" Then Draw_Spectrum
 
 On Error Resume Next
 If PlayerIsPlaying = "true" Then
   
    '// si se esta arrastrando el slider rep
    If bSlider = True Then Exit Sub
    
    If Stream_GetPosition(lCurrentChannel) > 10 And iCrossfadeTrack <> 0 Then iTimeCross = 5
    
    '// duracion de la rola
    If Stream_GetDuration(lCurrentChannel) - Stream_GetPosition(lCurrentChannel) <= iTimeCross Then
    
        '// si esta seleccionada el check para el loop
        If PlayerLoop = True Then Play: Exit Sub
     
        If frmPopUp.mnuOrdenAleatorio.Checked = True Then
          Randomize_Order
          Exit Sub
        End If
              
        frmPlayList.Next_Track
      Exit Sub
    End If
  
  '// Si esta la minimaskara
   If bMiniMask = True Then
     If bolTimeAct = False Then
       ScrollText(4).CaptionText = Convert_Time(Stream_GetPosition(lCurrentChannel))
     Else
       ScrollText(4).CaptionText = "-" & Convert_Time(Stream_GetDuration(lCurrentChannel) - Stream_GetPosition(lCurrentChannel))
     End If
     Slider(2).Value = CInt(Stream_GetPosition(lCurrentChannel))
   Else
     If bolTimeAct = False Then
       ScrollText(0).CaptionText = Convert_Time(Stream_GetPosition(lCurrentChannel))
     Else
       ScrollText(0).CaptionText = "-" & Convert_Time(Stream_GetDuration(lCurrentChannel) - Stream_GetPosition(lCurrentChannel))
     End If
     Slider(0).Value = CInt(Stream_GetPosition(lCurrentChannel))
   End If
       
   
  If tCurrentID3.Has_Lyrics3_Tag = True And bolLyricsShow = True Then
    Update_Lyrics
  End If
 End If

End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'| fUNCTION FOR ORDER THE LYRICS IN THE LIST FOR SHOW
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

 Sub Start_Lyrics()
  '// has lyrics
 LyricsIndex = 1
 
 If tCurrentID3.Has_Lyrics3_Tag = True And LyricsRef.ListCount > 0 Then
   '// form lyrics showing
   If bolLyricsShow = True Then
      frmLyrics.Reset_Values
      frmLyrics.lblArtist.Caption = tCurrentID3.Artist & " - " & tCurrentID3.Artist
      frmLyrics.lblTitle.Caption = tCurrentID3.Title
      frmLyrics.picLyrics.Visible = True
      frmLyrics.lblNoLyrics.Visible = False
   End If
 Else
   If bolLyricsShow = True Then
     frmLyrics.lblArtist.Caption = tCurrentID3.Artist & " - " & tCurrentID3.Album
     frmLyrics.lblTitle.Caption = tCurrentID3.Title
     frmLyrics.picLyrics.Visible = False
     frmLyrics.lblNoLyrics.Visible = True
   End If
 End If

End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Show_Lyrics(strLyrics As String)
 On Error Resume Next
   Dim RawLyrics() As String
   Dim Start As Integer
   Dim i As Integer, l As Integer, j As Integer, fin As Integer
   Dim strTemp As String
   Dim strTemp2 As String
   Dim startLyrics As Integer
   LyricsRef.Clear
  
   If Trim(strLyrics) = "" Then Exit Sub
   'check for timestamps
   If InStr(strLyrics, "[") = 0 Then Exit Sub
   'ok, it has lyrics, now put them into an array
   
   RawLyrics = Split(strLyrics, vbCr)
   l = UBound(RawLyrics)
   
   For i = 0 To l - 1
      Start = 1
      RawLyrics(i) = Trim(RawLyrics(i))
      Do
         j = InStr(Start, RawLyrics(i), "[")
         If j > 0 Then
            fin = InStr(Start, RawLyrics(i), "]")
            '// solo agregar letras hasta el formato 00:00:00
            If ((fin - 1) - j) < 9 Then
             '// extract time
              strTemp = Mid$(RawLyrics(i), j + 1, fin - j - 1)
              '// extract lyrics
               startLyrics = InStrRev(RawLyrics(i), "]", Len(RawLyrics(i)))
              strTemp2 = Right(RawLyrics(i), Len(RawLyrics(i)) - startLyrics)
              '// 00:00:00
              LyricsRef.AddItem strTemp & "    " & strTemp2
            End If
         End If
         Start = fin + 1
      Loop Until j = 0

   Next i
   
   If bolLyricsShow = True Then
      frmLyrics.Order_lblLyrics
   End If
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Update_Lyrics()
   Dim NumberOfLines As Integer
   Dim sCurrentTime As String
   Dim HH As Long, MM As Long, SS As Long
   Dim lTime As Long
   Dim tmp As String
    
   'now display the lyrics
   NumberOfLines = LyricsRef.ListCount
   On Error GoTo HELL
   
   lTime = Stream_GetPosition(lCurrentChannel)
   HH = lTime \ 3600      '// calkular horas
   MM = lTime \ 60 Mod 60 '// Calkular minutos
   SS = lTime Mod 60      '// calkular segundos
   SS = SS + 1
   If HH > 0 Then tmp = Format$(HH, "00:")
   sCurrentTime = tmp & Format$(MM, "00:") & Format$(SS, "00")

   
   'do we need to go forward?
   If sCurrentTime >= Trim(Left$(LyricsRef.List(LyricsIndex - 1), 9)) Then
      'yes.. how much??
      Do Until sCurrentTime <= Trim(Left$(LyricsRef.List(LyricsIndex), 9)) Or LyricsIndex = NumberOfLines
         frmLyrics.Move_Next_Focus_Lyrics
         LyricsIndex = LyricsIndex + 1
      Loop
   
   'do we need to go backwards?
   Else
      'yes, how much?

      Do Until sCurrentTime >= Trim(Left$(LyricsRef.List(LyricsIndex - 1), 9)) Or LyricsIndex = 1
         LyricsIndex = LyricsIndex - 1
         frmLyrics.Move_Previous_Focus_Lyrics
      Loop

   End If
   Exit Sub
HELL:

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Load_File_Tags()
 On Error Resume Next
 Dim sFullPath As String, sFileName As String, sFileEx As String
 Dim aFile() As String
   ' load tags
   tCurrentID3.Read_MPEGInfo = True
   tCurrentID3.Read_File_Tags sFileMainPlaying
   
   ScrollText(2).CaptionText = tCurrentID3.MPEG_Bit_Rate
   ScrollText(3).CaptionText = CStr(Left(tCurrentID3.MPEG_Frequency, 2))
   sFullPath = sFileMainPlaying
   aFile = Split(sFullPath, "\", , vbTextCompare)
   sFileName = aFile(UBound(aFile))
   sFileEx = Right(sFullPath, 3)
     
   sArtFile = tCurrentID3.Images
   sSysTrayText = tCurrentID3.Title & " - " & tCurrentID3.Artist & " - MMPlayerX"
   If OpcionesMusic.TaskBar = True = True Then frmPopUp.Caption = sSysTrayText
   If OpcionesMusic.SysTray = True Then CambiarIcono Text1.hwnd, Me.Icon.Handle, sSysTrayText

   sTextScroll = ""
   '// Song Name
   sTextScroll = Replace(sFormatScroll, "%S", tCurrentID3.Title)
   '// Artist
   sTextScroll = Replace(sTextScroll, "%A", tCurrentID3.Artist)
   '// Album
   sTextScroll = Replace(sTextScroll, "%B", tCurrentID3.Album)
   '// Year
   sTextScroll = Replace(sTextScroll, "%Y", tCurrentID3.Year)
   '// Genre
   sTextScroll = Replace(sTextScroll, "%G", tCurrentID3.Genre)
   '// Time
   sTextScroll = Replace(sTextScroll, "%T", tCurrentID3.MPEG_DurationTime)
   '// File Name
   sTextScroll = Replace(sTextScroll, "%N", sFileName)
   '// Time
   sTextScroll = Replace(sTextScroll, "%P", sFullPath)
   '// File extencion
   sTextScroll = Replace(sTextScroll, "%F", sFileEx)
  
   If bMiniMask = True Then
     ScrollText(5).CaptionText = sTextScroll
     ScrollText(5).ToolTipText = sTextScroll
   Else
     ScrollText(1).CaptionText = sTextScroll
     ScrollText(1).ToolTipText = sTextScroll
   End If

   '// karaoke function
   If tCurrentID3.Has_Lyrics3_Tag = True And Trim(tCurrentID3.Lyrics) <> "" Then
     Show_Lyrics Trim(tCurrentID3.Lyrics)
   Else
     LyricsRef.Clear
   End If
  
   '// Kolokar tooltiptext
   Show_ToolTipText
End Sub

Sub Show_ToolTipText()
' On Error Resume Next
  On Error GoTo HELL
  ttDemo.BackColor = frmPlayList.PL.BackColor + 1
  ttDemo.ForeColor = frmPlayList.PL.ForeColor
  ttDemo.TipText = "   Title: " & tCurrentID3.Title & vbCrLf & _
                   " Artist: " & tCurrentID3.Artist & vbCrLf & _
                   "Album: " & tCurrentID3.Album & vbCrLf & _
                   "  Year: " & tCurrentID3.Year & vbCrLf & _
                   "Genre: " & tCurrentID3.Genre
  If bMiniMask = True Then
     Set ttDemo.ParentControl = picMiniMode  ' ScrollText(5)
  Else
     Set ttDemo.ParentControl = picNormalMode
  End If
  ttDemo.Title = "MMPlayerX v 3.0"
  ttDemo.Icon = TTIconInfo
  ttDemo.Create
Exit Sub
HELL:
MsgBox err.Description
End Sub

Sub Form_Mover(X As Single, Y As Single)
 On Error Resume Next
  Dim DiffX As Long, DiffY As Long
  Dim NewX As Long, NewY As Long
  Dim ToLeftDistance As Long
  Dim ToRightDistance As Long
  Dim ToTopDistance As Long
  Dim ToBottomDistance As Long
  
    '// resta para mantener la posicion
    '// del cursor en la posicion inicial del objeto
        
    DiffX = X - StartDragX
    DiffY = Y - StartDragY
  
   If DiffX = 0 And DiffY = 0 Then Exit Sub
     '// obtener las coordenadas corectas
     NewX = Me.Left + DiffX
     NewY = Me.Top + DiffY
    
   If bHookForm = True Then
'       Me.Move NewX, NewY
       cAjustarDesk.UpdateDockDrag X, Y
          If mAttachedToBottom = True Then
             frmPlayList.Left = Me.Left + iPosHook
             frmPlayList.Top = frmMain.Top + frmMain.Height
             Exit Sub
          End If
          If mAttachedToTop = True Then
             frmPlayList.Left = Me.Left + iPosHook
             frmPlayList.Top = frmMain.Top - frmPlayList.Height
             Exit Sub
          End If
          If mAttachedToLeft = True Then
             frmPlayList.Top = Me.Top + iPosHook
             frmPlayList.Left = frmMain.Left - frmPlayList.Width
             Exit Sub
          End If
          
          If mAttachedToRight = True Then
             frmPlayList.Top = Me.Top + iPosHook
             frmPlayList.Left = frmMain.Left + frmMain.Width
             Exit Sub
          End If
   End If
    '// Enkontrar los bordes del escritorio
    
    ToRightDistance = frmPlayList.Left - (NewX + Me.Width)
    ToLeftDistance = NewX - (frmPlayList.Left + frmPlayList.Width)
    ToBottomDistance = frmPlayList.Top - (NewY + Me.Height)
    ToTopDistance = NewY - (frmPlayList.Top + frmPlayList.Height)
    '// si no esta anklado
    If Not mAttachedToBottom Then
        '// si esta en el area minima para arrastrarse para abajo
        If Abs(ToBottomDistance) <= mSnapDistance Then
           
           If (Me.Left + Me.Width >= frmPlayList.Left And Me.Left + Me.Width <= frmPlayList.Left + frmPlayList.Width) Or _
              (Me.Left >= frmPlayList.Left And Me.Left <= frmPlayList.Left + frmPlayList.Width) Then
                '// anklar al borde de abajo
                NewY = frmPlayList.Top - Me.Height
                mAttachedToBottom = True
                If Me.Left < frmPlayList.Left Then
                   iPosHook = Abs(Me.Left - frmPlayList.Left)
                Else
                   iPosHook = -Abs(Me.Left - frmPlayList.Left)
                End If
           End If

        End If
    Else
        
        If Abs(ToBottomDistance) > mSnapDistance Then
            '// Romper el anklado
            mAttachedToBottom = False
        Else
            '// mantener la actual posicion
            NewY = Me.Top
        End If
    End If

    If Not mAttachedToTop Then
        If Abs(ToTopDistance) <= mSnapDistance Then
           If (Me.Left + Me.Width >= frmPlayList.Left And Me.Left + Me.Width <= frmPlayList.Left + frmPlayList.Width) Or _
              (Me.Left >= frmPlayList.Left And Me.Left <= frmPlayList.Left + frmPlayList.Width) Then
              NewY = (frmPlayList.Top + frmPlayList.Height)
              mAttachedToTop = True
              If Me.Left < frmPlayList.Left Then
                   iPosHook = Abs(Me.Left - frmPlayList.Left)
              Else
                   iPosHook = -Abs(Me.Left - frmPlayList.Left)
              End If
           End If
        End If
    Else
        If Abs(ToTopDistance) > mSnapDistance Then
            mAttachedToTop = False
        Else
            NewY = Me.Top
        End If
    End If

    If Not mAttachedToRight Then
        If Abs(ToRightDistance) <= mSnapDistance Then
           If (Me.Top + Me.Height >= frmPlayList.Top And Me.Top + Me.Height <= frmPlayList.Top + frmPlayList.Height) Or _
              (Me.Top >= frmPlayList.Top And Me.Top <= frmPlayList.Top + frmPlayList.Height) Then

              NewX = frmPlayList.Left - Me.Width
              mAttachedToRight = True

              If Me.Top < frmPlayList.Top Then
                   iPosHook = Abs(Me.Top - frmPlayList.Top)
              Else
                   iPosHook = -Abs(Me.Top - frmPlayList.Top)
              End If
            End If
        End If
    Else
        If Abs(ToRightDistance) > mSnapDistance Then
            mAttachedToRight = False
        Else
            NewX = Me.Left
        End If
    End If

    If Not mAttachedToLeft Then
        If Abs(ToLeftDistance) <= mSnapDistance Then
           If (Me.Top + Me.Height >= frmPlayList.Top And Me.Top + Me.Height <= frmPlayList.Top + frmPlayList.Height) Or _
              (Me.Top >= frmPlayList.Top And Me.Top <= frmPlayList.Top + frmPlayList.Height) Then
              
              NewX = (frmPlayList.Left + frmPlayList.Width)
              mAttachedToLeft = True
              If Me.Top < frmPlayList.Top Then
                   iPosHook = Abs(Me.Top - frmPlayList.Top)
              Else
                   iPosHook = -Abs(Me.Top - frmPlayList.Top)
              End If
           End If
        End If
    Else
        If Abs(ToLeftDistance) > mSnapDistance Then
            mAttachedToLeft = False
        Else
            NewX = Me.Left
        End If
    End If
    If mAttachedToRight = True Or mAttachedToTop = True Or mAttachedToLeft = True Or mAttachedToBottom = True Then
      Me.Move NewX, NewY
      Exit Sub
    End If
   cAjustarDesk.UpdateDockDrag X, Y
   '// mover a la actual posicion
'   Me.Move NewX, NewY
End Sub


'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Form_Abajo(X As Single, Y As Single)
 On Error Resume Next
    
    '// variable para empezar a arrastrar
    bolDragMini = True
    '// almacenar las coordenadas iniciales
    StartDragX = X
    StartDragY = Y
     
End Sub

Sub Form_Arriba()
  If mAttachedToBottom = True Then bHookForm = True
  If mAttachedToTop = True Then bHookForm = True
  If mAttachedToLeft = True Then bHookForm = True
  If mAttachedToRight = True Then bHookForm = True
End Sub



