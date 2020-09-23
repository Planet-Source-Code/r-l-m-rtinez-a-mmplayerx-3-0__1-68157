VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTags 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " MPEG File Info Box + ID3 Tag Editor"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tags.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   1500
      Top             =   6210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ListView listRef 
      Height          =   1095
      Left            =   120
      TabIndex        =   34
      Top             =   6060
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   1931
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   17
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "FILE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "TRACK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "TITLE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ARTIST"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ALBUM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "YEAR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "GENRE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "COMMENTS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "COMPOSER"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ORG ARTIST"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   10
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "LINK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   11
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ENCODER"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   12
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "IMAGE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   13
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "LYRICS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   14
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "FILEPATH"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   15
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Edit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   16
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Row"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "Select All"
      Height          =   315
      Left            =   765
      TabIndex        =   0
      Top             =   5415
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   30
      TabIndex        =   31
      Top             =   -30
      Width           =   8475
      Begin ComctlLib.ProgressBar pbProgress 
         Height          =   270
         Left            =   60
         TabIndex        =   33
         Top             =   345
         Visible         =   0   'False
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblFile 
         Height          =   465
         Left            =   60
         TabIndex        =   32
         Top             =   135
         Width           =   8325
      End
   End
   Begin VB.Frame Frame 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5190
      Index           =   2
      Left            =   30
      TabIndex        =   30
      Top             =   615
      Width           =   3525
      Begin VB.ListBox fileTags 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4515
         Left            =   60
         MultiSelect     =   2  'Extended
         TabIndex        =   35
         Top             =   210
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6900
      TabIndex        =   21
      Top             =   5475
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5400
      TabIndex        =   20
      Top             =   5475
      Width           =   1305
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   3930
      TabIndex        =   19
      Top             =   5475
      Width           =   1305
   End
   Begin VB.PictureBox pictab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4230
      Index           =   1
      Left            =   3675
      ScaleHeight     =   4230
      ScaleWidth      =   4770
      TabIndex        =   22
      Top             =   1065
      Width           =   4770
      Begin VB.Frame Frame 
         Caption         =   "ID3"
         Height          =   4140
         Index           =   1
         Left            =   30
         TabIndex        =   23
         Top             =   45
         Width           =   4710
         Begin VB.TextBox txtEncoder 
            Height          =   315
            Left            =   1350
            TabIndex        =   51
            Top             =   3735
            Width           =   3270
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   8
            Left            =   1050
            TabIndex        =   50
            Top             =   3765
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.TextBox txtLink 
            Height          =   315
            Left            =   1350
            TabIndex        =   48
            Top             =   3405
            Width           =   3270
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   7
            Left            =   1050
            TabIndex        =   47
            Top             =   3435
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.TextBox txtOrgArtist 
            Height          =   315
            Left            =   1350
            TabIndex        =   45
            Top             =   3075
            Width           =   3270
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   6
            Left            =   1050
            TabIndex        =   44
            Top             =   3105
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.TextBox txtComposer 
            Height          =   315
            Left            =   1350
            TabIndex        =   42
            Top             =   2745
            Width           =   3270
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   1050
            TabIndex        =   41
            Top             =   2775
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.TextBox txtTrack 
            Height          =   315
            Left            =   1350
            MaxLength       =   3
            TabIndex        =   39
            Top             =   150
            Width           =   585
         End
         Begin VB.TextBox txtComment 
            Height          =   975
            Left            =   1350
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   1740
            Width           =   3285
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   1035
            TabIndex        =   36
            Top             =   1770
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   2535
            TabIndex        =   9
            Top             =   1455
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   1050
            TabIndex        =   7
            Top             =   1440
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   1050
            TabIndex        =   5
            Top             =   1140
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   1065
            TabIndex        =   3
            Top             =   840
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.ComboBox cboGenre 
            Height          =   315
            ItemData        =   "Tags.frx":000C
            Left            =   2820
            List            =   "Tags.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1425
            Width           =   1815
         End
         Begin VB.TextBox txtYear 
            Height          =   285
            Left            =   1350
            MaxLength       =   4
            TabIndex        =   8
            Top             =   1425
            Width           =   540
         End
         Begin VB.TextBox txtAlbum 
            Height          =   285
            Left            =   1350
            TabIndex        =   6
            Top             =   1125
            Width           =   3270
         End
         Begin VB.TextBox txtArtist 
            Height          =   315
            Left            =   1350
            TabIndex        =   4
            Top             =   810
            Width           =   3270
         End
         Begin VB.TextBox txtTitle 
            Height          =   315
            Left            =   1350
            TabIndex        =   2
            Top             =   480
            Width           =   3270
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Encoder:"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   52
            Top             =   3780
            Width           =   765
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Link:"
            Height          =   195
            Index           =   9
            Left            =   585
            TabIndex        =   49
            Top             =   3450
            Width           =   420
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Org. Artist:"
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   46
            Top             =   3120
            Width           =   960
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Composer:"
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   43
            Top             =   2790
            Width           =   960
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Track:"
            Height          =   195
            Index           =   2
            Left            =   495
            TabIndex        =   40
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Comment:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   37
            Top             =   1800
            Width           =   915
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Genre:"
            Height          =   195
            Index           =   7
            Left            =   1950
            TabIndex        =   28
            Top             =   1485
            Width           =   600
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Title:"
            Height          =   195
            Index           =   5
            Left            =   615
            TabIndex        =   27
            Top             =   510
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Album:"
            Height          =   195
            Index           =   4
            Left            =   435
            TabIndex        =   26
            Top             =   1155
            Width           =   615
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Artist:"
            Height          =   195
            Index           =   3
            Left            =   525
            TabIndex        =   25
            Top             =   855
            Width           =   525
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Year:"
            Height          =   195
            Index           =   1
            Left            =   585
            TabIndex        =   24
            Top             =   1455
            Width           =   465
         End
      End
   End
   Begin VB.PictureBox pictab 
      Appearance      =   0  'Flat
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
      Height          =   4230
      Index           =   2
      Left            =   3675
      ScaleHeight     =   4230
      ScaleWidth      =   4770
      TabIndex        =   53
      Top             =   1065
      Width           =   4770
      Begin VB.Frame Frame 
         Caption         =   "Art"
         Height          =   1890
         Index           =   3
         Left            =   210
         TabIndex        =   56
         Top             =   90
         Width           =   4380
         Begin VB.CheckBox chkTags 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   9
            Left            =   1335
            TabIndex        =   59
            Top             =   300
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.CommandButton cmdRemoveArt 
            Caption         =   "Remove Art"
            Height          =   345
            Left            =   105
            TabIndex        =   58
            Top             =   1155
            Width           =   1650
         End
         Begin VB.CommandButton cmdAddArt 
            Caption         =   "Add Art"
            Height          =   345
            Left            =   105
            TabIndex        =   57
            Top             =   750
            Width           =   1650
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Art Image:"
            Height          =   195
            Index           =   11
            Left            =   360
            TabIndex        =   60
            Top             =   315
            Width           =   945
         End
         Begin VB.Image imgArt 
            Height          =   1590
            Left            =   2190
            Stretch         =   -1  'True
            Top             =   210
            Width           =   1875
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "MPEG Info"
         Height          =   2145
         Index           =   0
         Left            =   195
         TabIndex        =   54
         Top             =   2010
         Width           =   4380
         Begin VB.Label lblMPEGInfo 
            Caption         =   "Label1"
            Height          =   1860
            Left            =   120
            TabIndex        =   55
            Top             =   210
            Width           =   4170
         End
      End
   End
   Begin VB.PictureBox pictab 
      Appearance      =   0  'Flat
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
      Height          =   4230
      Index           =   3
      Left            =   3675
      ScaleHeight     =   4230
      ScaleWidth      =   4770
      TabIndex        =   29
      Top             =   1065
      Width           =   4770
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Index           =   0
         Left            =   1485
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Index           =   1
         Left            =   1875
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Index           =   4
         Left            =   3030
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "||"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Index           =   2
         Left            =   2250
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPlayer 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Index           =   3
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   840
         TabIndex        =   17
         Top             =   3870
         Width           =   1425
      End
      Begin VB.CommandButton cmdUndo 
         Caption         =   "Deshacer"
         Height          =   315
         Left            =   2535
         TabIndex        =   18
         Top             =   3870
         Width           =   1425
      End
      Begin VB.TextBox txtLyrics 
         Height          =   3540
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   16
         Top             =   285
         Width           =   4695
      End
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   4740
      Left            =   3600
      TabIndex        =   1
      Top             =   675
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   8361
      TabFixedWidth   =   3175
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Tags         "
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "More        "
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Lyrics     "
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnnMusic As ADODB.Connection
Dim CMD As ADODB.Command

Private FilesSelected As Integer
 

'// vars functions undo in lyrics
Private Arr() As Long
Private Const cChunk = 10
Private Last As Long, Cur As Long
Dim Pos As Long

Dim FilePlaying As String
Dim LastPosition As Long
Dim LastState As String
Dim iCurrentAlbum As Integer
Dim bRestartPlayer As Boolean
Dim sLastGenre As String
Dim sArtFile As String

Private Sub cboGenre_Click()
 If sLastGenre = "" Then Exit Sub
 If sLastGenre <> cboGenre.Text Then
    Update_Tags_Ref
    If cmdApply.Enabled = False Then cmdApply.Enabled = True
 End If
 sLastGenre = ""
 
End Sub

Private Sub cboGenre_DropDown()
 sLastGenre = cboGenre.Text
End Sub

Private Sub chkTags_Click(Index As Integer)
  Dim bolEnabled As Boolean
  Dim bChk As Boolean
  Dim i As Integer
  
  If chkTags(Index).Value = vbChecked Then
    bolEnabled = True
  End If
  
  Select Case Index
    Case 0 '// Artist
       txtArtist.Enabled = bolEnabled
    Case 1 '// Album
       txtAlbum.Enabled = bolEnabled
    Case 2 '// Year
       txtYear.Enabled = bolEnabled
    Case 3 '// genre
       cboGenre.Enabled = bolEnabled
    Case 4 '// Comment
       txtComment.Enabled = bolEnabled
    Case 5 '// Composer
       txtComposer.Enabled = bolEnabled
    Case 6 '//Org Artis
       txtOrgArtist.Enabled = bolEnabled
    Case 7 '// Link
       txtLink.Enabled = bolEnabled
    Case 8 '// Encoder
       txtEncoder.Enabled = bolEnabled
    Case 9 '// Art image
       cmdAddArt.Enabled = bolEnabled
       cmdRemoveArt.Enabled = bolEnabled
  End Select
  
  For i = 0 To 9
    If chkTags(i).Value = vbChecked Then
       bChk = True
       Exit For
    End If
  Next
  
  cmdApply.Enabled = bChk
  
End Sub

Private Sub cmdAdd_Click()
   'add a timestamp at the beginning of the current line (Lyrics)
   
   Dim OldMin As Long 'the minutes of old timestamp
   Dim OldSec As Long 'the seconds of old timestamp
   Dim oldHou As Long 'the hours of old timestamp
   Dim NewMin As Long 'the minutes of new timestamp
   Dim NewSec As Long 'the seconds of new timestamp
   Dim NewHou As Long 'the hours of new timestamp
   Dim LineLength As Long 'length of a line
   Dim CurrentLine As Long 'the current line number
   Dim TotalLines As Long 'how many lines there are
   Dim sCurrentTime As String 'the current time in string format
   Dim CharPos As Long 'character position
   
   Dim arryOldTime() As String
   Dim arryNewTime() As String
   Dim s As String, strTemp As String
   Dim j As Integer, fin As Integer
   
    'error handler
   On Error GoTo HELL
     '================================================================
     '  This is simple lyrics function
     '  how it work? good question :)
     '   - First load a file in tag editor
     '   - Write the lyrics in the text
     '   - Play the song with the over buttons
     '   - Use add button in the just time
     '                            is all, ¿Facil no?
     '================================================================
   
   If fileTags.ListCount = 0 Or PlayerState = "false" Then Exit Sub
   
   'check to make sure it contains a time
   sCurrentTime = Convert_Time(Stream_GetPosition(1))
   arryNewTime = Split(sCurrentTime, ":")
   
   'if has hours
   If UBound(arryNewTime) > 1 Then
     'convert the Time into integers
     NewHou = Val(arryNewTime(0))
     NewMin = Val(arryNewTime(1))
     NewSec = Val(arryNewTime(2))
   Else
     NewHou = 0
     NewMin = Val(arryNewTime(0))
     NewSec = Val(arryNewTime(1))
   End If
   
   'add the brackets to the time
   s = "[" & sCurrentTime & "]"
   
   'set the insert point to the beginning of the line, add 1 to it to make sure
   'we don't get a 0 length string compare.
   CurrentLine = SendMessage(txtLyrics.hwnd, EM_LINEFROMCHAR, txtLyrics.SelStart, ZERO)
   CharPos = SendMessage(txtLyrics.hwnd, EM_LINEINDEX, CurrentLine, ZERO)
   'get the length of the line
   LineLength = SendMessage(txtLyrics.hwnd, EM_LINELENGTH, CharPos, ZERO)
   LineLength = CharPos + LineLength
   Pos = CharPos + 1
   
   '// note: the [Do..Loop Until] is optional for look only
   '// you can delete and work lyrics function :P
   
   'check to make sure there is no timestamp already there, if so
   'then compare the new time to the old timestamp so the new one
   'is inserted at the correct point in end of old timestamp.

      'there is a timestamp here, get the time
       Do
         j = InStr(Pos, txtLyrics.Text, "[")
         If j > 0 And j <= LineLength Then
            fin = InStr(Pos, txtLyrics.Text, "]")
            '// solo agregar letras hasta el formato 00:00:00
            If ((fin - 1) - j) < 9 Then
              strTemp = Mid$(txtLyrics.Text, j + 1, fin - j - 1)
            End If
         Else
           Exit Do
         End If
         
         arryOldTime = Split(strTemp, ":")
                
            'if has hours
          If UBound(arryOldTime) > 1 Then
             'convert the Time into integers
            oldHou = Val(arryOldTime(0))
            OldMin = Val(arryOldTime(1))
            OldSec = Val(arryOldTime(2))
          Else
            oldHou = 0
            OldMin = Val(arryOldTime(0))
            OldSec = Val(arryOldTime(1))
          End If
      
          'check to see if new timestamp is newer that old timestamp
          If (NewHou > oldHou) Or (NewHou = oldHou And NewMin > OldMin) Or (NewHou = oldHou And NewMin = OldMin And NewSec > OldSec) Then
             'yes, it is, so skip this one
             Pos = fin + 1
          Else
             Exit Do
          End If
       Loop Until j = 0
   LineLength = 0
    
   'subtract one from the insert point and insert the stamp
   Pos = Pos - 1
   txtLyrics.SelStart = Pos
   txtLyrics.SelText = s
   'and push this position onto the undo stack
    Undo_Push Pos
   'enable the undo button
   cmdUndo.Enabled = True
   
   'now drop them to the next non blank line, or back to the beginning
   'how many lines?
   TotalLines = SendMessage(txtLyrics.hwnd, EM_GETLINECOUNT, ZERO, ZERO)
   'safety check... should always be true
   If TotalLines > CurrentLine Then
      Do
         'increment current line
         CurrentLine = CurrentLine + 1
         'Get the position of the beginning of the line
         CharPos = SendMessage(txtLyrics.hwnd, EM_LINEINDEX, CurrentLine, ZERO)
         'get the length of the line
         LineLength = SendMessage(txtLyrics.hwnd, EM_LINELENGTH, CharPos, ZERO)
      'and keep looping until we get a non blank line or we get to the end
      Loop Until LineLength > 0 Or CurrentLine = TotalLines
      'if charpos = -1 then we are at the end.  Send them back to beginning
      If CharPos = -1 Then CharPos = 0
      'place cursor
      txtLyrics.SelStart = CharPos
   End If
   
   '/* update tags
   If Trim(txtLyrics.Text) <> "" Then Update_Tags_Ref
   
   'and set the focus back to the text box
   txtLyrics.SetFocus
   Exit Sub
HELL:

End Sub

Private Sub Save_Tags()
    
    Dim strFileName As String
    Dim i As Integer
    Dim iCount As Integer
    Dim iFUpdated As Integer
    Dim cID3 As New cMP3
    Dim bChk As Boolean
    Dim sSQL As String
    Dim iRow As Integer
    
    'On Error Resume Next
    On Error GoTo HELL
     '// if no checked all checkbox
    If fileTags.ListCount = 0 Then Exit Sub
    
    
    If FilesSelected > 1 Then
      For i = 0 To 9
        If chkTags(i).Value = vbChecked Then
            bChk = True
            Exit For
        End If
      Next

        If bChk = False Then
           Exit Sub
        End If
    End If
    
    '// reset values for progress bar
    pbProgress.min = 0
    pbProgress.Max = fileTags.ListCount
    pbProgress.Value = 0
    
    pbProgress.Visible = True
    cID3.Read_MPEGInfo = False
    For i = 0 To fileTags.ListCount - 1
        strFileName = listRef.ListItems(i + 1).Key
          
        DoEvents
          
       '// more than one files selected
       If FilesSelected > 1 Then
       
          If fileTags.Selected(i) = True Then
            lblFile.Caption = "Updating file: " & fileTags.List(i)
            cID3.Read_File_Tags strFileName
            '// make new tag
                       
            '// Artist checked change at all
            If chkTags(0).Value = vbChecked Then cID3.Artist = Trim(txtArtist.Text)

            '// Album checked change at all
            If chkTags(1).Value = vbChecked Then cID3.Album = Trim(txtAlbum.Text)

            '// year checked change at all
            If chkTags(2).Value = vbChecked Then cID3.Year = Trim(txtYear.Text)

            '// Genre checked change at all
            If chkTags(3).Value = vbChecked Then cID3.Genre = cboGenre.Text
            
            '// Comments checked change at all
            If chkTags(4).Value = vbChecked Then cID3.Comment = Trim(txtComment.Text)
            
            '// Composer checked change at all
            If chkTags(5).Value = vbChecked Then cID3.Composer = Trim(txtComposer.Text)
            
            '// Orig. Artist checked change at all
            If chkTags(6).Value = vbChecked Then cID3.OrigArtist = Trim(txtOrgArtist.Text)
            
            '// Links checked change at all
            If chkTags(7).Value = vbChecked Then cID3.LinkTo = Trim(txtLink.Text)
            
            '// Encoder checked change at all
            If chkTags(8).Value = vbChecked Then cID3.EncodedBy = Trim(txtEncoder.Text)
            
            '// Comments checked change at all
            If chkTags(9).Value = vbChecked Then cID3.Images = sArtFile
          
          'MODIFICAR LA LISTA DE REPRODUCCION CON LOS NUEVOS VALORES
          If Trim(listRef.ListItems(i + 1).SubItems(16)) <> "" Then
                iRow = CInt(listRef.ListItems(i + 1).SubItems(16))
                frmPlayList.PL.TextMatrix(iRow, 2) = cID3.Title
                frmPlayList.PL.TextMatrix(iRow, 3) = cID3.Artist
                frmPlayList.PL.TextMatrix(iRow, 4) = cID3.Album
                frmPlayList.PL.TextMatrix(iRow, 5) = cID3.Genre
                frmPlayList.PL.TextMatrix(iRow, 6) = cID3.Year
                frmPlayList.Agregar_PlayList_de_BaseDatos iRow
          End If
            
            '// write the tags
            cID3.Write_File_Tags

            If LCase(strFileName) = LCase(sFileMainPlaying) Then
               frmMain.Load_File_Tags
               If bolLyricsShow = True Then frmMain.Start_Lyrics
            End If

            iFUpdated = iFUpdated + 1
            
         End If
         
       ElseIf Trim(listRef.ListItems(i + 1).SubItems(15)) = "simon" Then
          cID3.Read_File_Tags strFileName
         '// make new tag
          cID3.TrackNr = Trim(listRef.ListItems.Item(i + 1).SubItems(1))
          cID3.Title = Trim(listRef.ListItems.Item(i + 1).SubItems(2))
          cID3.Artist = Trim(listRef.ListItems.Item(i + 1).SubItems(3))
          cID3.Album = Trim(listRef.ListItems.Item(i + 1).SubItems(4))
          cID3.Year = Trim(listRef.ListItems.Item(i + 1).SubItems(5))
          cID3.Genre = Trim(listRef.ListItems.Item(i + 1).SubItems(6))
          cID3.Comment = Trim(listRef.ListItems.Item(i + 1).SubItems(7))
          cID3.Composer = Trim(listRef.ListItems.Item(i + 1).SubItems(8))
          cID3.OrigArtist = Trim(listRef.ListItems.Item(i + 1).SubItems(9))
          cID3.LinkTo = Trim(listRef.ListItems.Item(i + 1).SubItems(10))
          cID3.EncodedBy = Trim(listRef.ListItems.Item(i + 1).SubItems(11))
          cID3.Images = Trim(listRef.ListItems.Item(i + 1).SubItems(12))
          cID3.Lyrics = Trim(listRef.ListItems.Item(i + 1).SubItems(13))
          
          '// write the tags
          cID3.Write_File_Tags
                              
'          sSQL = "UPDATE MUSIC SET TITLE='" & cID3.Title & "'," & _
'              "ARTIST='" & cID3.Artist & "',ALBUM='" & cID3.Album & "'," & _
'              "YEAR='" & cID3.Year & "',GENRE='" & cID3.Genre & "'" & _
'              " WHERE FILE='" & strFileName & "'"
'            Debug.Print sSQL
'           CMD.CommandText = sSQL
'           CMD.Execute
                                        
          If LCase(strFileName) = LCase(sFileMainPlaying) Then
             frmMain.Load_File_Tags
             If bolLyricsShow = True Then frmMain.Start_Lyrics
          End If
          
          listRef.ListItems.Item(i + 1).SubItems(1) = ""
          listRef.ListItems.Item(i + 1).SubItems(2) = ""
          listRef.ListItems.Item(i + 1).SubItems(3) = ""
          listRef.ListItems.Item(i + 1).SubItems(4) = ""
          listRef.ListItems.Item(i + 1).SubItems(5) = ""
          listRef.ListItems.Item(i + 1).SubItems(6) = ""
          listRef.ListItems.Item(i + 1).SubItems(7) = ""
          listRef.ListItems.Item(i + 1).SubItems(8) = ""
          listRef.ListItems.Item(i + 1).SubItems(9) = ""
          listRef.ListItems.Item(i + 1).SubItems(10) = ""
          listRef.ListItems.Item(i + 1).SubItems(11) = ""
          listRef.ListItems.Item(i + 1).SubItems(12) = ""
          listRef.ListItems.Item(i + 1).SubItems(13) = ""
          listRef.ListItems.Item(i + 1).SubItems(15) = ""
          
          'MODIFICAR LA LISTA DE REPRODUCCION CON LOS NUEVOS VALORES
          If Trim(listRef.ListItems(i + 1).SubItems(16)) <> "" Then
                iRow = CInt(listRef.ListItems(i + 1).SubItems(16))
                frmPlayList.PL.TextMatrix(iRow, 2) = cID3.Title
                frmPlayList.PL.TextMatrix(iRow, 3) = cID3.Artist
                frmPlayList.PL.TextMatrix(iRow, 4) = cID3.Album
                frmPlayList.PL.TextMatrix(iRow, 5) = cID3.Genre
                frmPlayList.PL.TextMatrix(iRow, 6) = cID3.Year
                frmPlayList.Agregar_PlayList_de_BaseDatos iRow
          End If
          
          iFUpdated = iFUpdated + 1
          
       End If
          
          iCount = iCount + 1
          pbProgress.Value = iCount


  Next
     pbProgress.Visible = False
     lblFile.Caption = " Listooooo! Updated [ " & iFUpdated & " ] files"
Exit Sub
HELL:
MsgBox err.Description

End Sub


Private Sub cmdAddArt_Click()
 Dialogo.Filter = "Image Files (*jpg, *.bmp, *.gif)|*.jpg;*.bmp;*.gif"
 Dialogo.ShowOpen

 If Dir(Dialogo.filename) = "" Or Dialogo.filename = "" Then Exit Sub
 sArtFile = Dialogo.filename
 imgArt.Picture = LoadPicture(sArtFile)
   Update_Tags_Ref

End Sub

Private Sub cmdApply_Click()
 On Error Resume Next
  cmdApply.Enabled = False
   Save_Tags
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub


Private Sub cmdOk_Click()
 cmdOk.Enabled = False
 If cmdApply.Enabled = True Then Save_Tags
 Unload Me
End Sub

Private Sub cmdPlayer_Click(Index As Integer)
  
  If fileTags.ListCount = 0 Then Exit Sub
  
  Select Case Index
    Case 0 '// skip backward
       Five_Seg_Backward
    Case 1 '// play
       If frmMain.PlayerIsPlaying = "true" Then frmMain.Stop_Player
       If fileTags.ListIndex = -1 Then fileTags.ListIndex = 0
       
       Player_Play listRef.ListItems(fileTags.ListIndex + 1).Key
    Case 2 '// pause
       Pause_Play
    Case 3 '// stop
       Stop_Player
       FilePlaying = ""
    Case 4 '// skip forward
       Five_Seg_Forward
  End Select
     txtLyrics.SetFocus

End Sub


Private Sub cmdRemoveArt_Click()
imgArt.Picture = LoadPicture()
sArtFile = ""
  Update_Tags_Ref

End Sub

Private Sub cmdSelAll_Click()
Dim i As Integer
For i = 0 To fileTags.ListCount - 1
 fileTags.Selected(i) = True
Next i
End Sub

Private Sub cmdUndo_Click()
 Dim fin As Integer, j As Integer, Start As Integer
  On Error GoTo HELL
  
    With txtLyrics
      Start = Undo_Pop
      If Start = 0 Then Start = 1
      'select the timestamp
       j = InStr(Start, txtLyrics.Text, "[")
         If j > 0 Then
            fin = InStr(Start + 1, txtLyrics.Text, "]")
            '// solo agregar letras hasta el formato 00:00:00
            If ((fin - 1) - j) > 9 Then
              fin = 0
            End If
         End If
      'get the postion of the last timestamp from the stack
      If Start = 1 Then Start = 0
      .SelStart = Start
      .SelLength = (fin - Start)
      'and delete it
      .SelText = ""
      .SetFocus
   End With
   'If there is nothing in the stack, undo should not be enabled
   If Cur = 0 Then cmdUndo.Enabled = False
Exit Sub
HELL:

End Sub


Private Sub Texts_Enableds(bolEnabled As Boolean)
   lblFile.Caption = ""
   chkTags(0).Value = vbUnchecked
   chkTags(1).Value = vbUnchecked
   chkTags(2).Value = vbUnchecked
   chkTags(3).Value = vbUnchecked
   chkTags(4).Value = vbUnchecked
   chkTags(5).Value = vbUnchecked
   chkTags(6).Value = vbUnchecked
   chkTags(7).Value = vbUnchecked
   chkTags(8).Value = vbUnchecked
   chkTags(9).Value = vbUnchecked
   
   
   chkTags(0).Enabled = Not bolEnabled
   chkTags(1).Enabled = Not bolEnabled
   chkTags(2).Enabled = Not bolEnabled
   chkTags(3).Enabled = Not bolEnabled
   chkTags(4).Enabled = Not bolEnabled
   chkTags(5).Enabled = Not bolEnabled
   chkTags(6).Enabled = Not bolEnabled
   chkTags(7).Enabled = Not bolEnabled
   chkTags(8).Enabled = Not bolEnabled
   chkTags(9).Enabled = Not bolEnabled
   
   chkTags(0).Visible = Not bolEnabled
   chkTags(1).Visible = Not bolEnabled
   chkTags(2).Visible = Not bolEnabled
   chkTags(3).Visible = Not bolEnabled
   chkTags(4).Visible = Not bolEnabled
   chkTags(5).Visible = Not bolEnabled
   chkTags(6).Visible = Not bolEnabled
   chkTags(7).Visible = Not bolEnabled
   chkTags(8).Visible = Not bolEnabled
   chkTags(9).Visible = Not bolEnabled
   
   txtTrack.Enabled = bolEnabled
   txtTitle.Enabled = bolEnabled
   txtArtist.Enabled = bolEnabled
   txtAlbum.Enabled = bolEnabled
   txtYear.Enabled = bolEnabled
   cboGenre.Enabled = bolEnabled
   txtComment.Enabled = bolEnabled
   txtComposer.Enabled = bolEnabled
   txtOrgArtist.Enabled = bolEnabled
   txtLink.Enabled = bolEnabled
   txtEncoder.Enabled = bolEnabled
   cmdAddArt.Enabled = bolEnabled
   cmdRemoveArt.Enabled = bolEnabled
   
   pictab(3).Enabled = bolEnabled
   txtLyrics.Text = ""
   lblMPEGInfo.Caption = ""
   

End Sub





Private Sub FileTags_Click()
 On Error Resume Next
 Dim i As Integer
 Dim tID3 As New cMP3
 FilesSelected = 0
 
 If fileTags.ListCount = 0 Then Exit Sub
 
 For i = 0 To fileTags.ListCount - 1
   If fileTags.Selected(i) = True Then
     FilesSelected = FilesSelected + 1
   End If
 Next i
 
 If PlayerState <> "false" Then Stop_Player
 
 
 '// pop for stack in lytics function
   Last = 10
   Cur = 0
   ReDim Arr(1 To Last) As Long
   cmdUndo.Enabled = False
 
 If FilesSelected > 1 Then
   Texts_Enableds False
   lblFile.Caption = LineLanguage(61)
   cmdApply.Enabled = False
   Exit Sub
 Else
   Texts_Enableds True
 End If
 
 lblFile.Caption = listRef.ListItems(fileTags.ListIndex + 1).Key
 lblFile.ToolTipText = listRef.ListItems(fileTags.ListIndex + 1).Key
 tID3.Read_MPEGInfo = True
 tID3.Read_File_Tags listRef.ListItems(fileTags.ListIndex + 1).Key
 
 txtTrack.Text = tID3.TrackNr
 txtTitle.Text = tID3.Title
 txtAlbum.Text = tID3.Album
 txtArtist.Text = tID3.Artist
 txtYear.Text = tID3.Year
 
 For i = 0 To cboGenre.ListCount - 1
     If cboGenre.List(i) = tID3.Genre Then
        cboGenre.ListIndex = i
        Exit For
     End If
 Next i
 
 txtComment.Text = tID3.Comment
 txtComposer.Text = tID3.Composer
 txtOrgArtist.Text = tID3.OrigArtist
 txtLink.Text = tID3.LinkTo
 If tID3.LinkTo = "" Then txtLink.Text = "www.geocities.com/skoria_36"
 txtEncoder.Text = tID3.EncodedBy
 If tID3.EncodedBy = "" Then txtEncoder.Text = "MMPlayerX Version 3.0"
 txtLyrics.Text = tID3.Lyrics
 imgArt.Picture = LoadPicture(tID3.Images)
 sArtFile = tID3.Images
 
  
 lblMPEGInfo.Caption = "<> Size: " & tID3.MPEG_FileSizeMB & "  <> Length: " & tID3.MPEG_DurationTime & vbCrLf & _
                       "<> MPEG " & tID3.MPEG_Version & vbCrLf & _
                       "<> Bitrate: " & tID3.MPEG_Bit_Rate & " kbps, " & IIf(tID3.MPEG_VBR, "variable bit rate", "constant bit rate") & vbCrLf & _
                       "<> " & tID3.MPEG_Frequency & " Hz  " & tID3.MPEG_ChannelMode & vbCrLf & _
                       "<> CRCs: " & tID3.MPEG_CRCs & "  <> Copyrighted: " & tID3.MPEG_Copyrighted & vbCrLf & _
                       "<> Original: " & tID3.MPEG_Original & "  <> Emphasis: " & tID3.MPEG_Emphasis & vbCrLf & _
                       "<> ID3 v1 tag: " & tID3.TagID3V1 & vbCrLf & _
                       "<> ID3 v2 tag: " & tID3.TagID3V2
Set tID3 = Nothing
End Sub

Private Sub Form_Load()
 On Error Resume Next
 Dim i As Integer
 Dim ID3 As New cMP3
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

 '// initialize values for undo functions
   Last = 10
   Cur = 0
   ReDim Arr(1 To Last) As Long
    
   For i = 0 To 147
        cboGenre.AddItem ID3.GetGenreName(i)
   Next i
  
  
  bolTagsShow = True
  
  
  Load_Language_Tags
  
  Me.Icon = frmMain.Icon
  frmTags.Left = (Screen.Width - frmTags.Width) / 2
  frmTags.Top = (Screen.Height - frmTags.Height) / 2

 PlayerState = "false"
 Set ID3 = Nothing
End Sub

Sub Load_Tags(sPath As String, Optional Row As Integer = -1)
 Dim sFileName As String
 Dim aFile() As String
 
 aFile = Split(sPath, "\", , vbTextCompare)
 sFileName = aFile(UBound(aFile))
 fileTags.AddItem sFileName
 Dim i As Integer
 
 If fileTags.ListCount = 0 Then Exit Sub
 listRef.ListItems.Add , sPath, sFileName
 listRef.ListItems.Item(sPath).SubItems(14) = sPath
 
 'SI LO ENVIAMOS DE LA LISTA DE REPRODUCCION
 If Row <> -1 Then
    listRef.ListItems.Item(sPath).SubItems(16) = Row
 End If
End Sub
Private Sub Player_Play(FilePlay As String)

On Error GoTo error
  If PlayerState = "pause" Then Pause_Play: Exit Sub
   
   Stream_Open FilePlay, FSOUND_NORMAL, 1, True, frmMain.VolumeNActuaL
   
   '// volume in main form
   'Stream_SetVolume 1, frmMain.VolumeNActuaL
   PlayerState = "true"
   FilePlaying = FilePlay
Exit Sub
error:
PlayerState = "false"
FilePlaying = ""
Stop_Player
End Sub

Sub Stop_Player()
 On Error Resume Next
  
 If fileTags.ListCount = 0 Then Exit Sub
 
 Stream_Stop 1
 PlayerState = "false"
End Sub

Private Sub Pause_Play()
 Dim CurState As Long
 Dim X
 
 On Error Resume Next
 
 If fileTags.ListCount = 0 Then Exit Sub
 
  If PlayerState = "false" Then Exit Sub
     CurState = Stream_GetState(1)
 '------'Esta Reproduciendo, pausar-------------------------------------------
     If CurState = 2 Then
       Stream_Pause 1
       PlayerState = "pause"
     Else
'------'Si esta pausado, reproducir---------------------------------------------
       Stream_Pause 1
       PlayerState = "true"
     End If
End Sub

Sub Five_Seg_Forward()
 On Error GoTo HELL
 Dim CurPos As Long
 
  If fileTags.ListCount = 0 Or PlayerState = "false" Then Exit Sub
  If PlayerState = "pause" Then Pause_Play
  
  CurPos = Stream_GetPosition(1)
  CurPos = CurPos + 5
  If CurPos > Stream_GetDuration(1) Then CurPos = Stream_GetDuration(1)
  Stream_SetPosition 1, CurPos
Exit Sub
HELL:

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Five_Seg_Backward()
 On Error GoTo HELL
 Dim CurPos As Long
  If fileTags.ListCount = 0 Or PlayerState = "false" Then Exit Sub
  If PlayerState = "pause" Then Pause_Play
  CurPos = Stream_GetPosition(1)
  CurPos = CurPos - 5
  If CurPos < 0 Then CurPos = 0
  Stream_SetPosition 1, CurPos
Exit Sub
HELL:
End Sub


Private Function Convert_Time(ByVal LSec As Long) As String
 Dim HH As Long, MM As Long, SS As Long
 Dim tmp As String
 
 HH = LSec \ 3600  '// calkular horas
 MM = LSec \ 60 Mod 60 '// Calkular minutos
 SS = LSec Mod 60  '// calkular segundos
 
 If HH > 0 Then tmp = Format$(HH, "00:")
 Convert_Time = tmp & Format$(MM, "00:") & Format$(SS, "00")
End Function


Private Sub Form_Unload(Cancel As Integer)
  If PlayerState <> "false" Then Stop_Player
  bolTagsShow = False
  cnnMusic.Close
  Set cnnMusic = Nothing

End Sub

Private Sub TabStrip_Click()
  pictab(TabStrip.SelectedItem.Index).ZOrder vbBringToFront

End Sub

'//------------------------------------------------------------------------------//
'// functions for undo function in lyrics
Private Sub Undo_Push(Arg As Long)
    Cur = Cur + 1
    On Error GoTo FailPush
        Arr(Cur) = Arg
    Exit Sub
FailPush:
    Last = Last + cChunk  ' Grow
    ReDim Preserve Arr(1 To Last) As Long
    Resume                  ' Try again
End Sub

Private Function Undo_Pop() As Long
    If Cur Then
        Undo_Pop = Arr(Cur)
        Cur = Cur - 1
        If Cur < (Last - cChunk) Then
            Last = Last - cChunk      ' Shrink
            ReDim Preserve Arr(1 To Last) As Long
        End If
    End If
End Function


Private Sub Update_Tags_Ref()
If fileTags.ListCount = 0 Then Exit Sub
If FilesSelected > 1 Then Exit Sub
  With listRef.ListItems.Item(fileTags.ListIndex + 1)
    .SubItems(1) = txtTrack.Text
    .SubItems(2) = txtTitle.Text
    .SubItems(3) = txtArtist.Text
    .SubItems(4) = txtAlbum.Text
    .SubItems(5) = txtYear.Text
    .SubItems(6) = cboGenre.List(cboGenre.ListIndex)
    .SubItems(7) = txtComment.Text
    .SubItems(8) = txtComposer.Text
    .SubItems(9) = txtOrgArtist.Text
    .SubItems(10) = txtLink.Text
    .SubItems(11) = txtEncoder.Text
    .SubItems(12) = sArtFile
    .SubItems(13) = txtLyrics.Text
    .SubItems(15) = "simon"
  End With
End Sub



Private Sub txtAlbum_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub

Private Sub txtArtist_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True
End Sub


Private Sub txtComment_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub

Private Sub txtLyrics_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
     KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
 If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub

Private Sub txtTitle_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True
End Sub

Private Sub txtYear_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub
Private Sub txtTrack_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub

Private Sub txtComposer_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub

Private Sub txtOrgArtist_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub

Private Sub txtLink_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub

Private Sub txtEncoder_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
   KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
  Update_Tags_Ref
  If cmdApply.Enabled = False Then cmdApply.Enabled = True

End Sub



