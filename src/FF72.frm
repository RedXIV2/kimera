VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PEditor 
   Caption         =   "P Editor"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   Icon            =   "FF72.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   532
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame PlaneOperationsFrame 
      Caption         =   "Plane Operations"
      Height          =   1575
      Left            =   1800
      TabIndex        =   84
      Top             =   6360
      Width           =   5655
      Begin VB.Frame PlaneGeometryFrame 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   2880
         TabIndex        =   91
         Top             =   120
         Width           =   2655
         Begin VB.CheckBox ShowPlaneCheck 
            Caption         =   "Show plane"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   0
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox XPlaneText 
            Height          =   285
            Left            =   240
            TabIndex        =   98
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox YPlaneText 
            Height          =   285
            Left            =   240
            TabIndex        =   97
            Text            =   "0"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox ZPlaneText 
            Height          =   285
            Left            =   240
            TabIndex        =   96
            Text            =   "0"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox AlphaPlaneText 
            Height          =   285
            Left            =   1800
            TabIndex        =   95
            Text            =   "0"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox BetaPlaneText 
            Height          =   285
            Left            =   1800
            TabIndex        =   94
            Text            =   "0"
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton InvertPlaneButton 
            Caption         =   "Invert"
            Height          =   255
            Left            =   1440
            TabIndex        =   93
            Top             =   0
            Width           =   615
         End
         Begin VB.CommandButton ResetPlaneButton 
            Caption         =   "Reset"
            Height          =   255
            Left            =   2040
            TabIndex        =   92
            Top             =   0
            Width           =   615
         End
         Begin MSComCtl2.UpDown ZPlaneUpDown 
            Height          =   285
            Left            =   840
            TabIndex        =   100
            Top             =   960
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "ZPlaneText"
            BuddyDispid     =   196614
            OrigLeft        =   720
            OrigTop         =   960
            OrigRight       =   960
            OrigBottom      =   1245
            Max             =   200
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown YPlaneUpDown 
            Height          =   285
            Left            =   840
            TabIndex        =   101
            Top             =   600
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "YPlaneText"
            BuddyDispid     =   196613
            OrigLeft        =   720
            OrigTop         =   600
            OrigRight       =   960
            OrigBottom      =   885
            Max             =   200
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown XPlaneUpDown 
            Height          =   285
            Left            =   840
            TabIndex        =   102
            Top             =   240
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "XPlaneText"
            BuddyDispid     =   196612
            OrigLeft        =   720
            OrigTop         =   240
            OrigRight       =   960
            OrigBottom      =   525
            Max             =   200
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown BetaPlaneUpDown 
            Height          =   285
            Left            =   2415
            TabIndex        =   103
            Top             =   840
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "BetaPlaneText"
            BuddyDispid     =   196616
            OrigLeft        =   5280
            OrigTop         =   840
            OrigRight       =   5520
            OrigBottom      =   1095
            Max             =   360
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown AlphaPlaneUpDown 
            Height          =   285
            Left            =   2400
            TabIndex        =   104
            Top             =   360
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "AlphaPlaneText"
            BuddyDispid     =   196615
            OrigLeft        =   5280
            OrigTop         =   480
            OrigRight       =   5520
            OrigBottom      =   735
            Max             =   360
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label BetaPlaneLabel 
            Caption         =   "Beta"
            Height          =   255
            Left            =   1200
            TabIndex        =   109
            Top             =   840
            Width           =   375
         End
         Begin VB.Label XPlaneLabel 
            Caption         =   "X"
            Height          =   255
            Left            =   0
            TabIndex        =   108
            Top             =   240
            Width           =   135
         End
         Begin VB.Label YPlaneLabel 
            Caption         =   "Y"
            Height          =   255
            Left            =   0
            TabIndex        =   107
            Top             =   600
            Width           =   135
         End
         Begin VB.Label ZPlaneLabel 
            Caption         =   "Z"
            Height          =   255
            Left            =   0
            TabIndex        =   106
            Top             =   960
            Width           =   135
         End
         Begin VB.Label AlphaPlaneLabel 
            Caption         =   "Alpha"
            Height          =   255
            Left            =   1200
            TabIndex        =   105
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton SlimButton 
         Caption         =   "Slim"
         Height          =   255
         Left            =   1440
         TabIndex        =   90
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton FlattenButton 
         Caption         =   "Fatten"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton CutModelButton 
         Caption         =   "Cut model"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton EraseLowerEmisphereButton 
         Caption         =   "Erase lower emisphere"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton MakeSymetricButton 
         Caption         =   "Make model simetric"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton MirrorHorizontallyButton 
         Caption         =   "Mirror model"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CheckBox ShowAxesCheck 
      Caption         =   "Show axes"
      Height          =   255
      Left            =   7560
      TabIndex        =   82
      Top             =   7680
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Frame GroupsFrame 
      Caption         =   "Groups"
      Height          =   1215
      Left            =   4560
      TabIndex        =   78
      Top             =   4080
      Width           =   2895
      Begin VB.CommandButton DownGroupCommand 
         Caption         =   "Down"
         Height          =   255
         Left            =   720
         TabIndex        =   111
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton UpGroupCommand 
         Caption         =   "Up"
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton HideShowGroupButton 
         Caption         =   "Hide/Show Group"
         Height          =   255
         Left            =   1320
         TabIndex        =   83
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton GroupPropertiesButton 
         Caption         =   "Group Properties"
         Height          =   255
         Left            =   1320
         TabIndex        =   81
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton DeleteGroupButton 
         Caption         =   "Delete Group"
         Height          =   255
         Left            =   1320
         TabIndex        =   80
         Top             =   240
         Width           =   1455
      End
      Begin VB.ListBox GroupsList 
         Height          =   645
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton SaveAsButton 
      Caption         =   "Save as"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton ApplyChangesButton 
      Caption         =   "Apply changes"
      Height          =   375
      Left            =   0
      TabIndex        =   67
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton LoadPButton 
      Caption         =   "Load p file"
      Height          =   375
      Left            =   0
      TabIndex        =   66
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame LightFrame 
      Caption         =   "Light"
      Height          =   1215
      Left            =   1800
      TabIndex        =   58
      Top             =   4080
      Width           =   2655
      Begin VB.CheckBox LightingCheck 
         Caption         =   "Enable lighting"
         Height          =   675
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   360
         Width           =   735
      End
      Begin VB.HScrollBar LightZScroll 
         Height          =   255
         Left            =   360
         Max             =   10
         Min             =   -10
         TabIndex        =   61
         Top             =   840
         Width           =   1335
      End
      Begin VB.HScrollBar LightYScroll 
         Height          =   255
         Left            =   360
         Max             =   10
         Min             =   -10
         TabIndex        =   60
         Top             =   600
         Width           =   1335
      End
      Begin VB.HScrollBar LightXScroll 
         Height          =   255
         Left            =   360
         Max             =   10
         Min             =   -10
         TabIndex        =   59
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Z:"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame MiscFrame 
      Caption         =   "Misc"
      Height          =   1095
      Left            =   1800
      TabIndex        =   54
      Top             =   5280
      Width           =   5655
      Begin VB.CommandButton KillLightingButton 
         Caption         =   "Kill precalculated lighting"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   5415
      End
      Begin VB.CommandButton DeletePolysColorCommand 
         Caption         =   "Delete all polygons with the selected color"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   5415
      End
      Begin VB.CommandButton DeletePolysNotColorCommand 
         Caption         =   "Delete all polygons but those with the selected color"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame RepositionFrame 
      Caption         =   "Reposition"
      Height          =   2055
      Left            =   7560
      TabIndex        =   44
      Top             =   2280
      Width           =   1935
      Begin VB.TextBox RepositionZText 
         Height          =   285
         Left            =   1320
         TabIndex        =   50
         Text            =   "0"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox RepositionYText 
         Height          =   285
         Left            =   1320
         TabIndex        =   49
         Text            =   "0"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox RepositionXText 
         Height          =   285
         Left            =   1320
         TabIndex        =   48
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.HScrollBar RepositionZ 
         Height          =   255
         Left            =   240
         Max             =   100
         Min             =   -100
         TabIndex        =   47
         Top             =   1680
         Width           =   975
      End
      Begin VB.HScrollBar RepositionY 
         Height          =   255
         Left            =   240
         Max             =   100
         Min             =   -100
         TabIndex        =   46
         Top             =   1080
         Width           =   975
      End
      Begin VB.HScrollBar RepositionX 
         Height          =   255
         Left            =   240
         Max             =   100
         Min             =   -100
         TabIndex        =   45
         Top             =   480
         Width           =   975
      End
      Begin VB.Label RepositionZLabel 
         Caption         =   "Z re-position"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label RepositionYLabel 
         Caption         =   "Y re-position"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label RepositionXLabel 
         Caption         =   "X re-position"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame RotateFrame 
      Caption         =   "Rotation"
      Height          =   2055
      Left            =   7560
      TabIndex        =   31
      Top             =   4440
      Width           =   1935
      Begin VB.HScrollBar RotateGamma 
         Height          =   255
         Left            =   240
         Max             =   360
         TabIndex        =   39
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox RotateGammaText 
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         Text            =   "0"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox RotateBetaText 
         Height          =   285
         Left            =   1320
         TabIndex        =   35
         Text            =   "0"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox RotateAlphaText 
         Height          =   285
         Left            =   1320
         TabIndex        =   34
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.HScrollBar RotateBeta 
         Height          =   255
         Left            =   240
         Max             =   360
         TabIndex        =   33
         Top             =   1080
         Width           =   975
      End
      Begin VB.HScrollBar RotateAlpha 
         Height          =   255
         Left            =   240
         Max             =   360
         TabIndex        =   32
         Top             =   480
         Width           =   975
      End
      Begin VB.Label RotateGammaLabel 
         Caption         =   "Gama rotation (Z-axis)"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label RotateBetaLabel 
         Caption         =   "Beta rotation (Y-axis)"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label RotateAlphaLabel 
         Caption         =   "Alpha rotation (X-axis)"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame ResizeFrame 
      Caption         =   "Resize"
      Height          =   2055
      Left            =   7560
      TabIndex        =   17
      Top             =   120
      Width           =   1935
      Begin VB.TextBox ResizeZText 
         Height          =   285
         Left            =   1320
         TabIndex        =   30
         Text            =   "100"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox ResizeYText 
         Height          =   285
         Left            =   1320
         TabIndex        =   29
         Text            =   "100"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox ResizeXText 
         Height          =   285
         Left            =   1320
         TabIndex        =   28
         Text            =   "100"
         Top             =   480
         Width           =   375
      End
      Begin VB.HScrollBar ResizeZ 
         Height          =   255
         Left            =   240
         Max             =   400
         TabIndex        =   24
         Top             =   1680
         Value           =   100
         Width           =   975
      End
      Begin VB.HScrollBar ResizeY 
         Height          =   255
         Left            =   240
         Max             =   400
         TabIndex        =   23
         Top             =   1080
         Value           =   100
         Width           =   975
      End
      Begin VB.HScrollBar ResizeX 
         Height          =   255
         Left            =   240
         Max             =   400
         TabIndex        =   22
         Top             =   480
         Value           =   100
         Width           =   975
      End
      Begin VB.Label ResizeZLabel 
         Caption         =   "Z re-size"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Y re-size"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label ResizeXLabel 
         Caption         =   "X re-size"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame ColorFrame 
      Caption         =   "Color editor"
      Height          =   5295
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
      Begin VB.CheckBox PalletizedCheck 
         Caption         =   "Palletized mode"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton LessBightnessButton 
         Caption         =   "-"
         Height          =   255
         Left            =   960
         TabIndex        =   43
         Top             =   4920
         Width           =   495
      End
      Begin VB.CommandButton MoreBightnessButton 
         Caption         =   "+"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   4920
         Width           =   495
      End
      Begin VB.TextBox ThersholdText 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   21
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox SelectedColorBText 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   20
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox SelectedColorGText 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   19
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox SelectedColorRText 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   18
         Top             =   2280
         Width           =   375
      End
      Begin VB.HScrollBar ThresholdSlider 
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   15
         Top             =   4320
         Value           =   20
         Width           =   975
      End
      Begin VB.HScrollBar SelectedColorB 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   10
         Top             =   3480
         Width           =   975
      End
      Begin VB.HScrollBar SelectedColorG 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   9
         Top             =   2880
         Width           =   975
      End
      Begin VB.HScrollBar SelectedColorR 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   8
         Top             =   2280
         Width           =   975
      End
      Begin VB.PictureBox PalletePicture 
         Height          =   1455
         Left            =   120
         ScaleHeight     =   93
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   93
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Brightness control"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Detection threshold"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Green level"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Blue level"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Red level"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Pallete:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame DrawModeFrame 
      Caption         =   "Draw Mode"
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
      Begin VB.OptionButton MeshOption 
         Caption         =   "Mesh"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton PolysOption 
         Caption         =   "Polygon colors"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton VertsOption 
         Caption         =   "Vetex colors"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox MainPicture 
      AutoSize        =   -1  'True
      Height          =   3975
      Left            =   1800
      MouseIcon       =   "FF72.frx":49E2
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   373
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.Frame CommandsFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   1215
      Left            =   7560
      TabIndex        =   68
      Top             =   6360
      Width           =   1935
      Begin VB.CommandButton NewPolyButton 
         Height          =   495
         Left            =   1440
         MaskColor       =   &H00FF0000&
         Picture         =   "FF72.frx":5824
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "New Polygon"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton PanButton 
         BackColor       =   &H0000C000&
         Height          =   495
         Left            =   960
         MaskColor       =   &H00FF0000&
         Picture         =   "FF72.frx":6466
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Panning"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton ZoomButton 
         BackColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         MaskColor       =   &H00FF0000&
         Picture         =   "FF72.frx":70A8
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Zoom In/Out"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton RotateButton 
         BackColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         MaskColor       =   &H00FF0000&
         Picture         =   "FF72.frx":7CEA
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Free Rotate"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton PickVertexButton 
         Height          =   495
         Left            =   1440
         MaskColor       =   &H00FF0000&
         Picture         =   "FF72.frx":892C
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Move vertex"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton EraseButton 
         Height          =   495
         Left            =   960
         MaskColor       =   &H00FF0000&
         Picture         =   "FF72.frx":956E
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Erase polygon"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton CutEdgeButton 
         Height          =   495
         Left            =   480
         MaskColor       =   &H00FF0000&
         Picture         =   "FF72.frx":A1B0
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Cut Edge"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton PaintButton 
         BackColor       =   &H80000016&
         Height          =   495
         Left            =   0
         MaskColor       =   &H00FF0000&
         Picture         =   "FF72.frx":ADF2
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Paint polygon"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
End
Attribute VB_Name = "PEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const K_PAINT = 0
Private Const K_CUT_EDGE = 1
Private Const K_ERASE_POLY = 2
Private Const K_PICK_VERTEX = 3
Private Const K_MOVE_VERTEX = 8
Private Const K_ROTATE = 4
Private Const K_ZOOM = 5
Private Const K_PAN = 6
Private Const K_NEW_POLY = 7

Private Const K_LOAD = -1
Private Const K_MOVE = 0
Private Const K_CLICK = 1
Private Const K_CLICK_SHIFT = 2

Private Const K_MESH = 0
Private Const K_PCOLORS = 1
Private Const K_VCOLORS = 2

Private Const K_PALLETIZED = 0
Private Const K_DIRECT_COLOR = 1

'Private Const UNDO_BUFFER_CAPACITY = 30

Private Const LETTER_SIZE = 5

Private Const LIGHT_STEPS = 10

Dim VColors_Original() As color
Dim DIST As Single
Dim mezclasDC As Long
Dim loaded As Boolean
Dim d_type As Byte
Dim color_table() As color
Dim translation_table_polys() As pair_i_b
Dim translation_table_vertex() As pair_i_b
Dim n_colors As Long
Dim threshold As Byte
Dim palDC As Long
Dim FullPalDC As Long
Dim Selected_color As Integer
Dim redX As Single, redY As Single, redZ As Single
Dim repX As Single, repY As Single, repZ As Single
Dim alpha As Single, Beta As Single, Gamma As Single
Dim x_last As Single, y_last As Integer
Dim rotate As Boolean
Dim file As String
Dim yl As Long
Dim direction As Boolean
Dim Select_mode As Byte
Dim cur_g As Integer
Dim v_buff(3) As Integer
Dim light_x As Integer, light_y As Integer, light_z As Integer
Dim Alpha_T As Single, Beta_T As Single
Dim tex_ids(0) As Long
Dim shifted As Integer
Dim VCountNewPoly As Integer
Dim VTempNewPoly(2) As Integer
Dim PrimaryFunction As Integer
Dim SecondaryFunction As Integer
Dim TernaryFunction As Integer
Dim PickedVertices() As Integer
Dim NumPickedVertices As Integer
Dim PickedVertexZ As Double
Dim AdjacentPolys() As Integer
Dim AdjacentVerts() As int_vector
Dim AdjacentAdjacentPolys() As int_vector
Dim ModelDirty As Boolean
Dim PanX As Single
Dim PanY As Single
Dim PanZ As Single

Dim PlaneOriginalVect1 As Point3D
Dim PlaneOriginalVect2 As Point3D
Dim PlaneOriginalPoint As Point3D

Dim PlaneOriginalPoint1 As Point3D
Dim PlaneOriginalPoint2 As Point3D
Dim PlaneOriginalPoint3 As Point3D
Dim PlaneOriginalPoint4 As Point3D

Dim PlaneRotationQuat As Quaternion
Dim PlaneTransformation(15) As Double

Dim PlaneVect1 As Point3D
Dim PlaneVect2 As Point3D
Dim PlanePoint As Point3D

Dim PlaneOriginalA As Single
Dim PlaneOriginalB As Single
Dim PlaneOriginalC As Single
Dim PlaneOriginalD As Single

Dim PlaneA As Single
Dim PlaneB As Single
Dim PlaneC As Single
Dim PlaneD As Single

Dim OldAlphaPlane As Single
Dim OldBetaPlane As Single
Dim OldGammaPlane As Single

Dim OGLContextEditor As Long
Dim GroupProperties As GroupPropertiesForm
Dim SelectedGroupIndex As Integer

Dim LoadingModifiersQ As Boolean
Dim SelectBoneForWeaponAttachmentQ As Boolean
Dim DoNotAddStateQ As Boolean

Dim ControlPressedQ As Boolean

Private Type PEditorState
    PanX As Single
    PanY As Single
    PanZ As Single

    DIST As Single

    redX As Single
    redY As Single
    redZ As Single

    repX As Single
    repY As Single
    repZ As Single

    RotateAlpha As Single
    RotateBeta As Single
    RotateGamma As Single

    alpha As Single
    Beta As Single
    Gamma As Single

    EditedPModel As PModel

    PalletizedQ As Boolean
    color_table() As color
    translation_table_polys() As pair_i_b
    translation_table_vertex() As pair_i_b
    n_colors As Long
    threshold As Byte
End Type

Dim UnDoBuffer() As PEditorState
Dim ReDoBuffer() As PEditorState

Dim UnDoCursor As Integer
Dim ReDoCursor As Integer

Private MinFormWidth As Integer
Private MinFormHeight As Integer

Private Function Min4(ByVal a As Long, ByVal B As Long, ByVal C As Long, ByVal d As Long)
    Dim i As Integer
    Dim temp(3) As Long

    temp(0) = B
    temp(1) = C
    temp(2) = d

    Min4 = a
    For i = 1 To 3
        If temp(i) < Min4 Then Min4 = temp(i)
    Next i
End Function

Private Sub AlphaPlaneText_Change()
    If IsNumeric(AlphaPlaneText.Text) Then
        If AlphaPlaneText.Text <= AlphaPlaneUpDown.max _
        And AlphaPlaneText.Text >= AlphaPlaneUpDown.Min Then
            AlphaPlaneUpDown.value = AlphaPlaneText.Text
        End If
    Else
        Beep
    End If
End Sub

Private Sub AlphaPlaneUpDown_Change()
    Dim diff As Single
    Dim aux_quat As Quaternion
    Dim res_quat As Quaternion

    diff = AlphaPlaneUpDown.value - OldAlphaPlane
    OldAlphaPlane = AlphaPlaneUpDown.value

    BuildQuaternionFromEuler diff, 0, 0, aux_quat
    MultiplyQuaternions PlaneRotationQuat, aux_quat, res_quat
    PlaneRotationQuat = res_quat
    BuildMatrixFromQuaternion PlaneRotationQuat, PlaneTransformation
    PlaneTransformation(12) = XPlaneUpDown.value * EditedPModel.diameter / 100
    PlaneTransformation(13) = YPlaneUpDown.value * EditedPModel.diameter / 100
    PlaneTransformation(14) = ZPlaneUpDown.value * EditedPModel.diameter / 100
    NormalizeQuaternion PlaneRotationQuat

    ComputeCurrentEquations

    MainPicture_Paint
End Sub

Private Sub ApplyChangesButton_Click()
    ApplyContextualizedPChanges False
    ModelEditor.UpdateEditedPiece
    SetOGLContext MainPicture.hdc, OGLContext
End Sub

Private Sub BetaPlaneText_Change()
    If IsNumeric(BetaPlaneText.Text) Then
        If BetaPlaneText.Text <= BetaPlaneUpDown.max _
        And BetaPlaneText.Text >= BetaPlaneUpDown.Min Then
            BetaPlaneUpDown.value = BetaPlaneText.Text
        End If
    Else
        Beep
    End If
End Sub

Private Sub BetaPlaneUpDown_Change()
    Dim diff As Single
    Dim aux_quat As Quaternion
    Dim res_quat As Quaternion

    diff = BetaPlaneUpDown.value - OldBetaPlane
    OldBetaPlane = BetaPlaneUpDown.value

    BuildQuaternionFromEuler 0, diff, 0, aux_quat
    MultiplyQuaternions PlaneRotationQuat, aux_quat, res_quat
    PlaneRotationQuat = res_quat
    BuildMatrixFromQuaternion PlaneRotationQuat, PlaneTransformation
    PlaneTransformation(12) = XPlaneUpDown.value * EditedPModel.diameter / 100
    PlaneTransformation(13) = YPlaneUpDown.value * EditedPModel.diameter / 100
    PlaneTransformation(14) = ZPlaneUpDown.value * EditedPModel.diameter / 100
    NormalizeQuaternion PlaneRotationQuat

    ComputeCurrentEquations

    MainPicture_Paint
End Sub

Private Sub CutModelButton_Click()
    Dim known_plane_pointsV() As Point3D
    If loaded Then
        AddStateToBuffer

        CutPModelThroughPlane EditedPModel, PlaneA, PlaneB, PlaneC, PlaneD, known_plane_pointsV

        If LightingCheck.value = 1 Then _
            ComputeNormals EditedPModel

        MainPicture_Paint
    End If
End Sub

Private Sub EraseLowerEmisphereButton_Click()
    Dim known_plane_pointsV() As Point3D
    If loaded Then
        AddStateToBuffer

        CutPModelThroughPlane EditedPModel, PlaneA, PlaneB, PlaneC, PlaneD, known_plane_pointsV
        EraseEmisphereVertices EditedPModel, PlaneA, PlaneB, PlaneC, PlaneD, False, known_plane_pointsV

        GroupProperties.Hide
        FillGroupsList

        If LightingCheck.value = 1 Then _
            ComputeNormals EditedPModel

        MainPicture_Paint
    End If

End Sub

Private Sub Form_Activate()
    SetOGLSettings
End Sub




Private Sub HideShowGroupButton_Click()
    If GroupsList.ListIndex >= 0 Then
        AddStateToBuffer

        GroupProperties.Hide
        EditedPModel.Groups(GroupsList.ListIndex).HiddenQ = _
            Not EditedPModel.Groups(GroupsList.ListIndex).HiddenQ

        FillGroupsList
        MainPicture_Paint
    End If
End Sub

Private Sub InvertPlaneButton_Click()
    AlphaPlaneText.Text = (AlphaPlaneText.Text + 180) Mod 360
    'BetaPlaneText.Text = (BetaPlaneText.Text + 180) Mod 360
End Sub

Private Sub LightingCheck_Click()
    If LightingCheck.value = vbChecked And loaded Then
        glEnable GL_NORMALIZE
        ComputeNormals EditedPModel
    Else
        glDisable GL_NORMALIZE
    End If

    MainPicture_Paint
End Sub

Private Sub CutEdgeButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PrimaryFunction = K_CUT_EDGE
    Else
        If Button = vbRightButton Then
            SecondaryFunction = K_CUT_EDGE
        Else
            TernaryFunction = K_CUT_EDGE
        End If
    End If
    SetFunctionButtonColors
End Sub

Private Sub DeleteGroupButton_Click()
    If GroupsList.ListCount > 1 Then
        If GroupsList.ListIndex >= 0 Then
            AddStateToBuffer

            GroupProperties.Hide
            VCountNewPoly = 0
            RemoveGroup EditedPModel, GroupsList.ListIndex
            CheckModelConsistency EditedPModel

            GroupProperties.Hide

            FillGroupsList
            MainPicture_Paint
        End If
    Else
        MsgBox "A P model must have at least one group", vbOKOnly, "Invalid operation"
    End If
End Sub

Private Sub DeletePolysColorCommand_Click()
    Dim pi_table As Integer
    Dim pi_model As Integer

    AddStateToBuffer
    pi_model = 0
    For pi_table = 0 To EditedPModel.head.NumPolys - 1
        If translation_table_polys(pi_table).i = Selected_color Then
            RemovePolygon EditedPModel, pi_model
        Else
            pi_model = pi_model + 1
        End If
    Next pi_table

    n_colors = 0
    fill_color_table EditedPModel, color_table, n_colors, translation_table_vertex, translation_table_polys, threshold

    MainPicture_Paint
End Sub

Private Sub DeletePolysNotColorCommand_Click()
    Dim pi_table As Integer
    Dim pi_model As Integer

    AddStateToBuffer
    pi_model = 0
    For pi_table = 0 To EditedPModel.head.NumPolys - 1
        If translation_table_polys(pi_table).i <> Selected_color Then
            RemovePolygon EditedPModel, pi_model
        Else
            pi_model = pi_model + 1
        End If
    Next pi_table

    n_colors = 0
    fill_color_table EditedPModel, color_table, n_colors, translation_table_vertex, translation_table_polys, threshold

    MainPicture_Paint
End Sub

Private Sub EraseButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PrimaryFunction = K_ERASE_POLY
    Else
        If Button = vbRightButton Then
            SecondaryFunction = K_ERASE_POLY
        Else
            TernaryFunction = K_ERASE_POLY
        End If
    End If
    SetFunctionButtonColors
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    DoNotAddStateQ = False

    If KeyCode = vbKeyHome Then _
        ResetCamera

    If KeyCode = vbKeyControl Then _
        ControlPressedQ = True

    If KeyCode = vbKeyZ And ControlPressedQ Then _
        UnDo
    If KeyCode = vbKeyY And ControlPressedQ Then _
        ReDo

    MainPicture_Paint
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyControl Then _
        ControlPressedQ = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload GroupProperties
End Sub

Private Sub GroupPropertiesButton_Click()
    If GroupsList.ListIndex >= 0 Then
        GroupProperties.SetSelectedGroup GroupsList.ListIndex
        GroupProperties.Show
    End If
End Sub

Private Sub LoadPButton_Click()
    Dim pattern As String
    Dim file As String

    On Error GoTo hand

    'MainPicture.Enabled = False
    pattern = "FF7 part file (field model)|*.p|FF7 part file (battle model)|*|3D Studio file|*.3ds"
    CommonDialog1.Filter = pattern
    CommonDialog1.ShowOpen 'Display the Open File Common Dialog

    file = EditedPModel.fileName
    If (CommonDialog1.fileName <> "") Then _
        OpenP CommonDialog1.fileName

    EditedPModel.fileName = file
    Timer1.Enabled = True
    Exit Sub
hand:
    If Err <> 32755 Then
        Dim mes As String
        mes = "Error" + Str(Err)
        MsgBox mes, vbOKOnly, "Unknow error Loading"
    End If
    Timer1.Enabled = True
End Sub

Private Sub MakeSymetricButton_Click()
    Dim known_plane_pointsV() As Point3D

    If loaded Then
        AddStateToBuffer
        CutPModelThroughPlane EditedPModel, PlaneA, PlaneB, PlaneC, PlaneD, known_plane_pointsV
        EraseEmisphereVertices EditedPModel, PlaneA, PlaneB, PlaneC, PlaneD, False, known_plane_pointsV
        DuplicateMirrorEmisphere EditedPModel, PlaneA, PlaneB, PlaneC, PlaneD
        CheckModelConsistency EditedPModel

        GroupProperties.Hide
        FillGroupsList

        If LightingCheck.value = 1 Then _
            ComputeNormals EditedPModel

        MainPicture_Paint
    End If
End Sub

Private Sub ResetPlaneButton_Click()
    ResetPlane

    MainPicture_Paint
End Sub

Private Sub SaveAsButton_Click()
    Dim pattern As String
    Dim file As String

    On Error GoTo hand

    MainPicture.Enabled = False
    pattern = "FF7 part file (field model)|*.p|FF7 part file (battle model)|*"
    CommonDialog1.Filter = pattern
    CommonDialog1.ShowSave 'Display the Open File Common Dialog

    If (CommonDialog1.fileName <> "") Then
        SaveP CommonDialog1.fileName
        ResizeX.value = 100
        ResizeY.value = 100
        ResizeZ.value = 100
        repX = 0
        repY = 0
        repZ = 0
        RotateAlpha.value = 0
        RotateBeta.value = 0
        RotateGamma.value = 0
        SetOGLContext MainPicture.hdc, OGLContext
    End If
    Timer1.Enabled = True
    Exit Sub
hand:
    If Err <> 32755 Then
        Dim mes As String
        mes = "Error" + Str(Err)
        MsgBox mes, vbOKOnly, "Unknow error Saving"
    End If
    Timer1.Enabled = True
End Sub

Private Sub ShowAxesCheck_Click()
    Call MainPicture_Paint
End Sub

Private Sub ShowPlaneCheck_Click()
    MainPicture_Paint
End Sub

Private Sub SlimButton_Click()
Dim i As Integer
    Dim trans_inverse(15) As Double
    Dim min_x As Single, min_y As Single, min_z As Single
    Dim max_x As Single, max_y As Single, max_z As Single
    If loaded Then
        AddStateToBuffer

        With EditedPModel.BoundingBox
            min_x = .min_x
            min_y = .min_y
            min_z = .min_z
            max_x = .max_x
            max_y = .max_y
            max_z = .max_z
        End With

        For i = 0 To 15
            trans_inverse(i) = PlaneTransformation(i)
        Next i

        InvertMatrix trans_inverse
        ApplyPModelTransformation EditedPModel, trans_inverse
        ComputeBoundingBox EditedPModel

        Slim EditedPModel
        ApplyPModelTransformation EditedPModel, PlaneTransformation
        With EditedPModel.BoundingBox
            .min_x = min_x
            .min_y = min_y
            .min_z = min_z
            .max_x = max_x
            .max_y = max_y
            .max_z = max_z
        End With

        If LightingCheck.value = 1 Then ComputeNormals EditedPModel
        MainPicture_Paint
    End If
End Sub

Private Sub KillLightingButton_Click()
    If loaded Then
        AddStateToBuffer

        KillPrecalculatedLighting EditedPModel, translation_table_vertex
        ApplyColorTable EditedPModel, color_table, translation_table_vertex
        MainPicture_Paint
    End If
End Sub
Private Sub KillTexButton_Click()
    RemoveTexturedGroups EditedPModel
    MainPicture_Paint
End Sub
Private Sub NewPolyButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PrimaryFunction = K_NEW_POLY
    Else
        If Button = vbRightButton Then
            SecondaryFunction = K_NEW_POLY
        Else
            TernaryFunction = K_NEW_POLY
        End If
    End If
    SetFunctionButtonColors
End Sub

Private Sub PaintButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PrimaryFunction = K_PAINT
    Else
        If Button = vbRightButton Then
            SecondaryFunction = K_PAINT
        Else
            TernaryFunction = K_PAINT
        End If
    End If
    SetFunctionButtonColors
End Sub

Private Sub PalletizedCheck_Click()
    If PalletizedCheck.value = 0 Then
        ThresholdSlider.Enabled = 0
        If Not ModelDirty Then _
            CopyVColors VColors_Original, EditedPModel.vcolors
        ModelDirty = False
        Selected_color = -1
        DeletePolysNotColorCommand.Enabled = False
        DeletePolysColorCommand.Enabled = False
        KillLightingButton.Enabled = False
    Else
        ThresholdSlider.Enabled = 1
        CopyVColors EditedPModel.vcolors, VColors_Original
        n_colors = 0
        fill_color_table EditedPModel, color_table, n_colors, translation_table_vertex, translation_table_polys, threshold
        ApplyColorTable EditedPModel, color_table, translation_table_vertex
        n_colors = n_colors - 1
        DeletePolysNotColorCommand.Enabled = True
        DeletePolysColorCommand.Enabled = True
        KillLightingButton.Enabled = True
    End If
    DrawPallete K_CLICK
    ModelDirty = False
End Sub

Private Sub PanButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PrimaryFunction = K_PAN
    Else
        If Button = vbRightButton Then
            SecondaryFunction = K_PAN
        Else
            TernaryFunction = K_PAN
        End If
    End If
    SetFunctionButtonColors
End Sub

Private Sub PickVertexButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
        PrimaryFunction = K_PICK_VERTEX
    Else
        If Button = vbRightButton Then
            SecondaryFunction = K_PICK_VERTEX
        Else
            TernaryFunction = K_PICK_VERTEX
        End If
    End If
    SetFunctionButtonColors
End Sub


Private Sub PalletePicture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim s_row As Single
    Dim n_rows As Long
    Dim xc, yc As Integer
    Dim col As Long

    If loaded = True And Button <> 0 Then
        If PalletizedCheck.value Then
            s_row = 2 * PalletePicture.ScaleHeight / n_colors
            n_rows = PalletePicture.ScaleHeight / s_row
            yc = Fix(y / PalletePicture.ScaleHeight * n_rows)
            If x > PalletePicture.ScaleWidth / 2 Then
                xc = 1
            Else
                xc = 0
            End If

            Selected_color = yc * 2 + xc

            SelectedColorR.value = color_table(Selected_color).r
            SelectedColorG.value = color_table(Selected_color).g
            SelectedColorB.value = color_table(Selected_color).B
            DrawPallete K_MOVE
        Else
            col = GetPixel(PalletePicture.hdc, x, y)
            SelectedColorR.value = getRed(col)
            SelectedColorG.value = getGreen(col)
            SelectedColorB.value = getBlue(col)
            DrawPallete K_MOVE
        End If
    End If
End Sub
Private Sub PalletePicture_Paint()
    DrawPallete K_CLICK
End Sub

Private Sub RotateButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PrimaryFunction = K_ROTATE
    Else
        If Button = vbRightButton Then
            SecondaryFunction = K_ROTATE
        Else
            TernaryFunction = K_ROTATE
        End If
    End If
    SetFunctionButtonColors
End Sub
Private Sub MirrorHorizontallyButton_Click()
    If loaded Then
        AddStateToBuffer

        MirrorEmisphere EditedPModel, PlaneA, PlaneB, PlaneC, PlaneD

        If LightingCheck.value = 1 Then _
            ComputeNormals EditedPModel

        MainPicture_Paint
    End If
End Sub
Private Sub MoreBightnessButton_Click()
    Dim i As Integer

    If loaded Then
        AddStateToBuffer

        ChangeBrigthness EditedPModel, 5

        MainPicture_Paint
    End If
End Sub

Private Sub LessBightnessButton_Click()
    Dim i As Integer

    If loaded Then
        AddStateToBuffer

        ChangeBrigthness EditedPModel, -5

        MainPicture_Paint
    End If
End Sub
Private Sub FlattenButton_Click()
    Dim i As Integer
    Dim trans_inverse(15) As Double
    Dim min_x As Single, min_y As Single, min_z As Single
    Dim max_x As Single, max_y As Single, max_z As Single
    If loaded Then
        AddStateToBuffer

        With EditedPModel.BoundingBox
            min_x = .min_x
            min_y = .min_y
            min_z = .min_z
            max_x = .max_x
            max_y = .max_y
            max_z = .max_z
        End With

        For i = 0 To 15
            trans_inverse(i) = PlaneTransformation(i)
        Next i

        InvertMatrix trans_inverse
        ApplyPModelTransformation EditedPModel, trans_inverse
        ComputeBoundingBox EditedPModel

        Fatten EditedPModel
        ApplyPModelTransformation EditedPModel, PlaneTransformation
        With EditedPModel.BoundingBox
            .min_x = min_x
            .min_y = min_y
            .min_z = min_z
            .max_x = max_x
            .max_y = max_y
            .max_z = max_z
        End With

        If LightingCheck.value = 1 Then ComputeNormals EditedPModel
        MainPicture_Paint
    End If
End Sub
Private Sub Form_Load()
    ReDim UnDoBuffer(UNDO_BUFFER_CAPACITY)
    ReDim ReDoBuffer(UNDO_BUFFER_CAPACITY)

    LightXScroll.max = LIGHT_STEPS
    LightXScroll.Min = -LIGHT_STEPS
    LightYScroll.max = LIGHT_STEPS
    LightYScroll.Min = -LIGHT_STEPS
    LightZScroll.max = LIGHT_STEPS
    LightZScroll.Min = -LIGHT_STEPS

    SelectedColorR.Enabled = False
    SelectedColorG.Enabled = False
    SelectedColorB.Enabled = False
    ThresholdSlider.Enabled = False

    RotateGamma.Enabled = False
    RotateBeta.Enabled = False
    RotateAlpha.Enabled = False
    RepositionX.Enabled = False
    RepositionY.Enabled = False
    RepositionZ.Enabled = False

    ResizeZ.Enabled = False
    ResizeY.Enabled = False
    ResizeX.Enabled = False

    SelectedColorRText.Enabled = False
    SelectedColorGText.Enabled = False
    SelectedColorBText.Enabled = False
    ThersholdText.Enabled = False

    ResizeXText.Enabled = False
    ResizeYText.Enabled = False
    ResizeZText.Enabled = False

    RepositionZText.Enabled = False
    RepositionYText.Enabled = False
    RepositionXText.Enabled = False

    RotateAlphaText.Enabled = False
    RotateBetaText.Enabled = False
    RotateGammaText.Enabled = False

    d_type = 2
    loaded = False
    rotate = False
    Selected_color = -1
    threshold = 60
    alpha = 0
    Beta = 0
    Gamma = 0
    mezclasDC = creaDC(MainPicture.hdc, MainPicture.ScaleWidth, MainPicture.ScaleHeight)
    palDC = creaDC(PalletePicture.hdc, PalletePicture.ScaleWidth, PalletePicture.ScaleHeight)
    FullPalDC = creaDC(PalletePicture.hdc, PalletePicture.ScaleWidth, PalletePicture.ScaleHeight)
    direction = True
    Select_mode = 0
    LightingCheck.value = 0
    PrimaryFunction = K_ROTATE
    SecondaryFunction = K_ZOOM
    TernaryFunction = K_PAN
    DrawPallete K_LOAD
    Set GroupProperties = Forms.Add("GroupPropertiesForm")
    OGLContextEditor = CreateOGLContext(MainPicture.hdc)

    MinFormWidth = Me.width
    MinFormHeight = Me.height
End Sub
Private Sub Form_Paint()
    Me.Caption = "P Editor - " + file
End Sub

Private Sub Form_Resize()
    'The window can't be resize while minimized
    If Me.WindowState = 1 Then _
        Exit Sub
    If Me.width < MinFormWidth Then _
        Me.width = MinFormWidth
    If Me.height < MinFormHeight Then _
        Me.height = MinFormHeight
    If Me.ScaleWidth > 0 Then
        Me.ResizeFrame.Left = Me.ScaleWidth - Me.ResizeFrame.width
        Me.RepositionFrame.Left = Me.ScaleWidth - Me.RepositionFrame.width
        Me.RotateFrame.Left = Me.ScaleWidth - Me.RotateFrame.width
        CommandsFrame.Left = Me.RotateFrame.Left
        ShowAxesCheck.Left = Me.RotateFrame.Left

        Me.MainPicture.width = Me.ScaleWidth - Me.ResizeFrame.width - Me.ColorFrame.width - _
                                (Me.MainPicture.Left - (Me.ColorFrame.Left + Me.ColorFrame.width)) _
                                - 5
        Me.MainPicture.height = Me.ScaleHeight - Me.LightFrame.height - Me.MiscFrame.height _
                                - Me.PlaneOperationsFrame.height - Me.MainPicture.Top

        'Central frames placement
        Me.LightFrame.width = Me.MainPicture.width / 2
        Me.GroupsFrame.width = Me.MainPicture.width / 2
        Me.GroupsFrame.Left = Me.LightFrame.Left + Me.LightFrame.width
        Me.MiscFrame.width = Me.MainPicture.width
        Me.PlaneOperationsFrame.width = Me.MiscFrame.width
        Me.LightFrame.Top = Me.MainPicture.Top + Me.MainPicture.height
        Me.GroupsFrame.Top = Me.LightFrame.Top
        Me.MiscFrame.Top = Me.LightFrame.Top + Me.LightFrame.height
        Me.PlaneOperationsFrame.Top = Me.MiscFrame.Top + Me.MiscFrame.height

        Me.PlaneGeometryFrame.Left = Screen.TwipsPerPixelX * (Me.PlaneOperationsFrame.width - 3) _
                                    - Me.PlaneGeometryFrame.width

        'Plane operations buttons placement
        Me.MirrorHorizontallyButton.width = Me.PlaneGeometryFrame.Left _
                                            - Me.MirrorHorizontallyButton.Left * 2
        Me.MakeSymetricButton.width = Me.MirrorHorizontallyButton.width
        Me.EraseLowerEmisphereButton.width = Me.MirrorHorizontallyButton.width
        Me.CutModelButton.width = Me.MirrorHorizontallyButton.width
        Me.FlattenButton.width = Me.MirrorHorizontallyButton.width / 2
        Me.SlimButton.width = Me.FlattenButton.width
        Me.SlimButton.Left = Me.FlattenButton.width + Me.FlattenButton.Left

        'Misc operations buttons placement
        Me.DeletePolysNotColorCommand.width = Me.MiscFrame.width * Screen.TwipsPerPixelX _
                                                - Me.DeletePolysNotColorCommand.Left * 2
        Me.DeletePolysColorCommand.width = Me.DeletePolysNotColorCommand.width
        Me.KillLightingButton.width = Me.DeletePolysNotColorCommand.width


        Me.DeleteGroupButton.width = (GroupsFrame.width / 2 + 2) * Screen.TwipsPerPixelX
        Me.DeleteGroupButton.Left = (GroupsFrame.width / 2 - 5) * Screen.TwipsPerPixelX

        Me.GroupPropertiesButton.width = Me.DeleteGroupButton.width
        Me.GroupPropertiesButton.Left = Me.DeleteGroupButton.Left

        Me.HideShowGroupButton.width = Me.DeleteGroupButton.width
        Me.HideShowGroupButton.Left = Me.DeleteGroupButton.Left

        Me.GroupsList.width = (GroupsFrame.width / 2 - 15) * Screen.TwipsPerPixelX

        MainPicture_Paint
    End If
End Sub

Private Sub Form_Terminate()
    DeleteDC mezclasDC
    DeleteDC palDC

    DisableOpenGL OGLContextEditor
    Unload GroupProperties
End Sub

Private Sub SelectedColorR_Change()
    If Selected_color > -1 Then
        If Not DoNotAddStateQ Then AddStateToBuffer
        DoNotAddStateQ = True

        color_table(Selected_color).r = SelectedColorR.value
        ApplyColorTable EditedPModel, color_table, translation_table_vertex
        CopyVColors EditedPModel.vcolors, VColors_Original
        ModelDirty = True
        MainPicture_Paint
        DoNotAddStateQ = False
    End If
    SelectedColorRText.Text = SelectedColorR.value
    DrawPallete K_MOVE
End Sub

Private Sub LightXScroll_Change()
    light_x = LightXScroll.value
    Call MainPicture_Paint
End Sub

Private Sub LightYScroll_Change()
    light_y = LightYScroll.value
    Call MainPicture_Paint
End Sub

Private Sub LightZScroll_Change()
    light_z = LightZScroll.value
    Call MainPicture_Paint
End Sub

Private Sub SelectedColorG_Change()
    If Selected_color > -1 Then
        If Not DoNotAddStateQ Then AddStateToBuffer
        DoNotAddStateQ = True

        color_table(Selected_color).g = SelectedColorG.value
        ApplyColorTable EditedPModel, color_table, translation_table_vertex
        CopyVColors EditedPModel.vcolors, VColors_Original
        ModelDirty = True
        Call MainPicture_Paint
        DoNotAddStateQ = False
    End If
    DrawPallete K_MOVE
    SelectedColorGText.Text = SelectedColorG.value
End Sub
Private Sub SelectedColorB_Change()
    If Selected_color > -1 Then
        If Not DoNotAddStateQ Then AddStateToBuffer
        DoNotAddStateQ = True

        color_table(Selected_color).B = SelectedColorB.value

        ApplyColorTable EditedPModel, color_table, translation_table_vertex
        CopyVColors EditedPModel.vcolors, VColors_Original
        ModelDirty = True
        Call MainPicture_Paint
        DoNotAddStateQ = False
    End If
    DrawPallete K_MOVE
    SelectedColorBText.Text = SelectedColorB.value
End Sub

Private Sub ThresholdSlider_Change()
    Dim i As Integer
    threshold = ThresholdSlider.value
    n_colors = 0

    If Not ModelDirty Then
        CopyVColors VColors_Original, EditedPModel.vcolors
    Else
        If Not DoNotAddStateQ Then AddStateToBuffer
        DoNotAddStateQ = True

        CopyVColors EditedPModel.vcolors, VColors_Original
        ModelDirty = False
    End If

    fill_color_table EditedPModel, color_table, n_colors, translation_table_vertex, translation_table_polys, threshold
    ApplyColorTable EditedPModel, color_table, translation_table_vertex

    MainPicture_Paint
    PalletePicture_Paint
    ThersholdText.Text = ThresholdSlider.value
    DoNotAddStateQ = False
End Sub
Private Sub Label13_Click()
    Beta = RotateBeta.value
    RotateBetaText.Text = RotateBeta.value
    Call MainPicture_Paint
End Sub
Private Sub MeshOption_Click()
    d_type = 0
    MainPicture_Paint
End Sub
Private Sub PolysOption_Click()
    d_type = 1
    ComputePColors EditedPModel
    MainPicture_Paint
End Sub


Private Sub Timer1_Timer()
    Sleep 100
    Timer1.Enabled = False
    MainPicture.Enabled = True
End Sub



Private Sub VertsOption_Click()
    d_type = 2
    MainPicture_Paint
End Sub
Private Sub MainPicture_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If loaded Then
        SetOGLSettings

        If LightingCheck.value = vbChecked Then glEnable GL_LIGHTING

        glClearColor 0.5, 0.5, 1, 0
        glViewport 0, 0, MainPicture.ScaleWidth, MainPicture.ScaleHeight
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT

        SetCameraPModel EditedPModel, PanX, PanY, _
                        PanZ + DIST, alpha, _
                        Beta, Gamma, 1, 1, 1

        ConcatenateCameraModelViewQuat repX, repY, repZ, _
            EditedPModel.RotationQuaternion, redX, redY, redZ

        x_last = x
        y_last = y

        If Button = vbLeftButton Then
            DoFunction PrimaryFunction, K_CLICK + Shift, x, y
        Else
            If Button = vbRightButton Then
                DoFunction SecondaryFunction, K_CLICK + Shift, x, y
            Else
                If Button = vbMiddleButton Then _
                    DoFunction TernaryFunction, K_CLICK + Shift, x, y
            End If
        End If

        MainPicture_Paint
    End If
End Sub

Private Sub MainPicture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If loaded And Button <> 0 Then
        glClearColor 0.5, 0.5, 1, 0
        glViewport 0, 0, MainPicture.ScaleWidth, MainPicture.ScaleHeight
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT

        SetCameraPModel EditedPModel, PanX, PanY, _
                        PanZ + DIST, alpha, _
                        Beta, Gamma, 1, 1, 1

        ConcatenateCameraModelView repX, repY, repZ, _
            RotateAlpha.value, RotateBeta.value, RotateGamma.value, redX, redY, redZ

        If Button = vbLeftButton Then
            DoFunction PrimaryFunction, K_MOVE, x, y
        Else
            If Button = vbRightButton Then
                DoFunction SecondaryFunction, K_MOVE, x, y
            Else
                If Button = vbMiddleButton Then _
                    DoFunction TernaryFunction, K_MOVE, x, y
            End If
        End If

        x_last = x
        y_last = y

        MainPicture_Paint
    End If
End Sub

Private Sub MainPicture_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        If Button = vbLeftButton Then
            If PrimaryFunction = K_MOVE_VERTEX Then _
                PrimaryFunction = K_PICK_VERTEX
        Else
            If SecondaryFunction = K_MOVE_VERTEX Then _
                SecondaryFunction = K_PICK_VERTEX
        End If
End Sub

Private Sub PalletePicture_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim s_row As Single
    Dim n_rows As Long
    Dim xc, yc As Integer
    Dim col As Long

    If loaded = True Then
        If PalletizedCheck.value Then
            s_row = 2 * PalletePicture.ScaleHeight / n_colors
            n_rows = PalletePicture.ScaleHeight / s_row
            yc = Fix(y / PalletePicture.ScaleHeight * n_rows)
            If x > PalletePicture.ScaleWidth / 2 Then
                xc = 1
            Else
                xc = 0
            End If

            Selected_color = yc * 2 + xc

            SelectedColorR.value = color_table(Selected_color).r
            SelectedColorG.value = color_table(Selected_color).g
            SelectedColorB.value = color_table(Selected_color).B
            DrawPallete K_CLICK
        Else
            col = GetPixel(PalletePicture.hdc, x, y)
            SelectedColorR.value = getRed(col)
            SelectedColorG.value = getGreen(col)
            SelectedColorB.value = getBlue(col)
            DrawPallete K_CLICK
        End If
    End If

End Sub

Private Sub MainPicture_Paint()
    Dim p_min As Point3D
    Dim p_max As Point3D

    Dim model_diameter_normalized As Single

    Dim vi As Integer
    If loaded = True Then
        glViewport 0, 0, MainPicture.ScaleWidth, MainPicture.ScaleHeight
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT

        SetCameraPModel EditedPModel, PanX, PanY, _
                        PanZ + DIST, alpha, _
                        Beta, Gamma, 1, 1, 1

        glMatrixMode GL_MODELVIEW
        glPushMatrix
        ConcatenateCameraModelView repX, repY, repZ, _
            RotateAlpha.value, RotateBeta.value, RotateGamma.value, redX, redY, redZ

        If LightingCheck.value = 1 Then
            glDisable GL_LIGHT0
            glDisable GL_LIGHT1
            glDisable GL_LIGHT2
            glDisable GL_LIGHT3
            ComputePModelBoundingBox EditedPModel, p_min, p_max
            model_diameter_normalized = (-2 * ComputeSceneRadius(p_min, p_max)) / LIGHT_STEPS
            SetLighting GL_LIGHT0, model_diameter_normalized * LightXScroll.value, _
                                    model_diameter_normalized * LightYScroll.value, _
                                    model_diameter_normalized * LightZScroll.value, 1, 1, 1, False
        Else
            glDisable GL_LIGHTING
        End If

        SetDefaultOGLRenderState

        Select Case (d_type)
            Case K_MESH:
                DrawPModelMesh EditedPModel
            Case K_PCOLORS:
                glEnable GL_POLYGON_OFFSET_FILL
                glPolygonOffset 1, 1
                DrawPModelPolys EditedPModel
                glDisable GL_POLYGON_OFFSET_FILL
                DrawPModelMesh EditedPModel
            Case K_VCOLORS:
                DrawPModel EditedPModel, tex_ids(), True
        End Select

        SetDefaultOGLRenderState
        If ShowPlaneCheck.value = vbChecked Then _
            DrawPlane

        If ShowAxesCheck.value = vbChecked Then _
            DrawAxes

        SwapBuffers MainPicture.hdc
    End If
End Sub

Private Sub SelectedColorRText_Change()
    If IsNumeric(SelectedColorRText.Text) Then
        If SelectedColorRText.Text >= 0 And SelectedColorRText.Text <= 255 Then
            SelectedColorR.value = SelectedColorRText.Text
            Call SelectedColorR_Change
        Else
            SelectedColorRText.Text = 255
            Call SelectedColorR_Change
        End If
    Else
        Beep
        SelectedColorRText.Text = 0
        Call SelectedColorR_Change
    End If
End Sub
Private Sub SelectedColorGText_Change()
    If IsNumeric(SelectedColorGText.Text) Then
        If SelectedColorGText.Text >= 0 And SelectedColorGText.Text <= 255 Then
            SelectedColorG.value = SelectedColorGText.Text
            Call SelectedColorG_Change
        Else
            SelectedColorGText.Text = 255
            Call SelectedColorG_Change
        End If
    Else
        Beep
        SelectedColorGText.Text = 0
        Call SelectedColorG_Change
    End If
End Sub
Private Sub SelectedColorBText_Change()
    If IsNumeric(SelectedColorBText.Text) Then
        If SelectedColorBText.Text >= 0 And SelectedColorBText.Text <= 255 Then
            SelectedColorB.value = SelectedColorBText.Text
            SelectedColorB_Change
        Else
            SelectedColorBText.Text = 255
            SelectedColorB_Change
        End If
    Else
        Beep
        SelectedColorBText.Text = 0
        SelectedColorB_Change
    End If
End Sub
Private Sub ThersholdText_Change()
    If IsNumeric(ThersholdText.Text) Then
        If ThersholdText.Text >= 0 And ThersholdText.Text <= 255 Then
            ThresholdSlider.value = ThersholdText.Text
            ThresholdSlider_Change
        Else
            ThersholdText.Text = 255
            ThresholdSlider_Change
        End If
    Else
        Beep
        ThersholdText.Text = 0
        ThresholdSlider_Change
    End If
End Sub


Private Sub XPlaneText_Change()
    If IsNumeric(XPlaneText.Text) Then
        If XPlaneText.Text <= XPlaneUpDown.max _
        And XPlaneText.Text >= XPlaneUpDown.Min Then
            XPlaneUpDown.value = XPlaneText.Text
        End If
    Else
        Beep
    End If
End Sub

Private Sub XPlaneUpDown_Change()
    PlaneTransformation(12) = XPlaneUpDown.value * EditedPModel.diameter / 100

    ComputeCurrentEquations

    MainPicture_Paint
End Sub

Private Sub YPlaneText_Change()
   If IsNumeric(YPlaneText.Text) Then
        If YPlaneText.Text <= YPlaneUpDown.max _
        And YPlaneText.Text >= YPlaneUpDown.Min Then
            YPlaneUpDown.value = YPlaneText.Text
        End If
    Else
        Beep
    End If
End Sub

Private Sub YPlaneUpDown_Change()
    PlaneTransformation(13) = YPlaneUpDown.value * EditedPModel.diameter / 100

    ComputeCurrentEquations

    MainPicture_Paint
End Sub


Private Sub ZoomButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PrimaryFunction = K_ZOOM
    Else
        If Button = vbRightButton Then
            SecondaryFunction = K_ZOOM
        Else
            TernaryFunction = K_ZOOM
        End If
    End If
    SetFunctionButtonColors
End Sub
Private Sub SetFunctionButtonColors()
    SetPaintButtonColor
    SetCutEdgeButtonColor
    SetEraseButtonColor
    SetPickVertexButtonColor
    SetRotateButtonColor
    SetZoomButtonColor
    SetPanButtonColor
    SetNewPolyButtonColor
End Sub
Private Sub SetPaintButtonColor()
    If PrimaryFunction = K_PAINT Then
        PaintButton.BackColor = vbRed
    Else
        If SecondaryFunction = K_PAINT Then
            PaintButton.BackColor = vbBlue
        Else
            If TernaryFunction = K_PAINT Then
                PaintButton.BackColor = vbGreen
            Else
                PaintButton.BackColor = &H80000016
            End If
        End If
    End If
End Sub
Private Sub SetCutEdgeButtonColor()
    If PrimaryFunction = K_CUT_EDGE Then
        CutEdgeButton.BackColor = vbRed
        d_type = K_PCOLORS
        PolysOption.value = True
    Else
        If SecondaryFunction = K_CUT_EDGE Then
            CutEdgeButton.BackColor = vbBlue
            d_type = K_PCOLORS
            PolysOption.value = True
        Else
            If TernaryFunction = K_CUT_EDGE Then
                CutEdgeButton.BackColor = vbGreen
                d_type = K_PCOLORS
                PolysOption.value = True
            Else
                CutEdgeButton.BackColor = &H80000016
            End If
        End If
    End If
End Sub
Private Sub SetEraseButtonColor()
    If PrimaryFunction = K_ERASE_POLY Then
        EraseButton.BackColor = vbRed
    Else
        If SecondaryFunction = K_ERASE_POLY Then
            EraseButton.BackColor = vbBlue
        Else
            If TernaryFunction = K_ERASE_POLY Then
                EraseButton.BackColor = vbGreen
            Else
                EraseButton.BackColor = &H80000016
            End If
        End If
    End If
End Sub
Private Sub SetPickVertexButtonColor()
    If PrimaryFunction = K_PICK_VERTEX Then
        PickVertexButton.BackColor = vbRed
    Else
        If SecondaryFunction = K_PICK_VERTEX Then
            PickVertexButton.BackColor = vbBlue
        Else
            If TernaryFunction = K_PICK_VERTEX Then
                PickVertexButton.BackColor = vbGreen
            Else
                PickVertexButton.BackColor = &H80000016
            End If
        End If
    End If
End Sub
Private Sub SetRotateButtonColor()
    If PrimaryFunction = K_ROTATE Then
        RotateButton.BackColor = vbRed
    Else
        If SecondaryFunction = K_ROTATE Then
            RotateButton.BackColor = vbBlue
        Else
            If TernaryFunction = K_ROTATE Then
                RotateButton.BackColor = vbGreen
            Else
                RotateButton.BackColor = &H80000016
            End If
        End If
    End If
End Sub
Private Sub SetZoomButtonColor()
    If PrimaryFunction = K_ZOOM Then
        ZoomButton.BackColor = vbRed
    Else
        If SecondaryFunction = K_ZOOM Then
            ZoomButton.BackColor = vbBlue
        Else
            If TernaryFunction = K_ZOOM Then
                ZoomButton.BackColor = vbGreen
            Else
                ZoomButton.BackColor = &H80000016
            End If
        End If
    End If
End Sub
Private Sub SetPanButtonColor()
    If PrimaryFunction = K_PAN Then
        PanButton.BackColor = vbRed
    Else
        If SecondaryFunction = K_PAN Then
            PanButton.BackColor = vbBlue
        Else
            If TernaryFunction = K_PAN Then
                PanButton.BackColor = vbGreen
            Else
                PanButton.BackColor = &H80000016
            End If
        End If
    End If
End Sub
Private Sub SetNewPolyButtonColor()
    If PrimaryFunction = K_NEW_POLY Then
        NewPolyButton.BackColor = vbRed
    Else
        If SecondaryFunction = K_NEW_POLY Then
            NewPolyButton.BackColor = vbBlue
        Else
            If TernaryFunction = K_NEW_POLY Then
                NewPolyButton.BackColor = vbGreen
            Else
                NewPolyButton.BackColor = &H80000016
            End If
        End If
    End If
    VCountNewPoly = 0
End Sub
Private Sub DoFunction(ByVal NFunc As Integer, ByVal Ev As Integer, ByVal x As Integer, ByVal y As Integer)
    Dim Al As Single

    Dim PI As Integer
    Dim ei As Integer
    Dim vi As Integer
    Dim vi1 As Integer
    Dim vi2 As Integer
    Dim g_index As Integer
    Dim g_index_aux As Integer
    Dim max_p As Integer

    Dim c_temp As color

    Dim p_temp As Point3D
    Dim p_temp2 As Point3D

    Dim p1 As Point3D
    Dim p2 As Point3D

    Dim tc1 As Point2D
    Dim tc2 As Point2D

    Dim intersection_point As Point3D
    Dim intersection_tex_coord As Point2D

    Select Case NFunc
        Case K_PAINT:
            'Change polygon color/get polygon color
            If Ev >= K_CLICK Then
                PI = GetClosestPolygon(EditedPModel, x, y, repZ + PanZ + DIST)

                If PI > -1 Then
                    AddStateToBuffer

                    If PalletizedCheck.value = 1 Then
                        If Ev = K_CLICK_SHIFT Then
                            Selected_color = translation_table_vertex(EditedPModel.polys(PI).Verts(0) + EditedPModel.Groups(GetPolygonGroup(EditedPModel.Groups, PI)).offvert).i
                            With color_table(Selected_color)
                                SelectedColorR.value = .r
                                SelectedColorG.value = .g
                                SelectedColorB.value = .B
                            End With
                        Else
                            If Selected_color > -1 Then
                               PaintPolygon EditedPModel, PI, SelectedColorR.value, _
                                                        SelectedColorG.value, _
                                                        SelectedColorB.value

                               UpdateTranslationTable translation_table_vertex, EditedPModel, PI, Selected_color
                               ModelDirty = True
                            End If
                        End If
                    Else
                        If Ev = K_CLICK_SHIFT Then
                            c_temp = ComputePolyColor(EditedPModel, PI)
                            With c_temp
                                SelectedColorR.value = .r
                                SelectedColorG.value = .g
                                SelectedColorB.value = .B
                            End With
                        Else
                            PaintPolygon EditedPModel, PI, SelectedColorR.value, _
                                                     SelectedColorG.value, _
                                                     SelectedColorB.value
                        End If
                    End If

                   ' If glIsEnabled(GL_LIGHTING) = GL_TRUE Then _
                   '     ComputeNormals EditedPModel
                End If
            End If
        Case K_CUT_EDGE:
            'Cut an edge on the clicked point (thus slpitting the surrounding polygons)
            If Ev = K_CLICK Then
                PI = GetClosestPolygon(EditedPModel, x, y, repZ + PanZ + DIST)

                If PI > -1 Then
                    AddStateToBuffer

                    ei = GetClosestEdge(EditedPModel, PI, x, y, Al)

                    vi1 = EditedPModel.polys(PI).Verts(ei)
                    vi2 = EditedPModel.polys(PI).Verts((ei + 1) Mod 3)
                    g_index = GetPolygonGroup(EditedPModel.Groups, PI)
                    p1 = EditedPModel.Verts(EditedPModel.Groups(g_index).offvert + vi1)
                    p2 = EditedPModel.Verts(EditedPModel.Groups(g_index).offvert + vi2)

                    If EditedPModel.Groups(g_index).texFlag = 1 Then
                        tc1 = EditedPModel.TexCoords(EditedPModel.Groups(g_index).offTex + vi1)
                        tc2 = EditedPModel.TexCoords(EditedPModel.Groups(g_index).offTex + vi2)
                    End If

                    intersection_point = GetPointInLine(p1, p2, Al)
                    intersection_tex_coord = GetPointInLine2D(tc1, tc2, Al)

                    CutEdgeAtPoint EditedPModel, PI, ei, intersection_point, intersection_tex_coord

                    While FindNextAdjacentPolyEdge(EditedPModel, p1, p2, PI, ei)
                        'If crossed group boundaries, recompute intersection_tex_coord
                        g_index_aux = GetPolygonGroup(EditedPModel.Groups, PI)
                        If (g_index_aux <> g_index) Then
                            g_index = g_index_aux
                            If EditedPModel.Groups(g_index).texFlag = 1 Then
                                vi1 = EditedPModel.polys(PI).Verts(ei)
                                vi2 = EditedPModel.polys(PI).Verts((ei + 1) Mod 3)
                                tc1 = EditedPModel.TexCoords(EditedPModel.Groups(g_index).offTex + vi1)
                                tc2 = EditedPModel.TexCoords(EditedPModel.Groups(g_index).offTex + vi2)
                                intersection_tex_coord = GetPointInLine2D(tc1, tc2, Al)
                            End If
                        End If
                        g_index = g_index_aux
                        CutEdgeAtPoint EditedPModel, PI, ei, intersection_point, intersection_tex_coord
                    Wend

                    ComputePColors EditedPModel
                    If glIsEnabled(GL_LIGHTING) = GL_TRUE Then _
                        ComputeNormals EditedPModel

                    If PalletizedCheck.value = 1 Then
                        n_colors = 0
                        fill_color_table EditedPModel, color_table, n_colors, translation_table_vertex, _
                                         translation_table_polys, threshold
                    End If
                End If
            End If
        Case K_ERASE_POLY:
            'Erase polygon
            If Ev = K_CLICK Then
                PI = GetClosestPolygon(EditedPModel, x, y, repZ + PanZ + DIST)
                If PI > -1 Then
                    AddStateToBuffer

                    If EditedPModel.head.NumPolys > 1 Then
                        RemovePolygon EditedPModel, PI
                    Else
                        MsgBox "A P model must have at least one polygon. This last triangle can't be removed"
                    End If

                    'If glIsEnabled(GL_LIGHTING) = GL_TRUE Then _
                    '   ComputeNormals EditedPModel

                    If PalletizedCheck.value = 1 Then
                        n_colors = 0
                        fill_color_table EditedPModel, color_table, n_colors, translation_table_vertex, _
                                         translation_table_polys, threshold
                    End If
                End If
            End If
        Case K_PICK_VERTEX:
            'Pick a vertex. When a vertex is picked, switch to the K_MOVE_VERTEX operation
            If Ev = K_CLICK Then
                PI = GetClosestVertex(EditedPModel, x, y, repZ + PanZ + DIST)
                If PI > -1 Then
                    AddStateToBuffer

                    PickedVertexZ = GetVertexProjectedDepth(EditedPModel.Verts, PI)

                    NumPickedVertices = GetEqualVertices(EditedPModel, PI, PickedVertices)
                    If PrimaryFunction = K_PICK_VERTEX Then
                        PrimaryFunction = K_MOVE_VERTEX
                    Else
                        SecondaryFunction = K_MOVE_VERTEX
                    End If

                    If glIsEnabled(GL_LIGHTING) = GL_TRUE Then
                        GetAllNormalDependentPolys EditedPModel, PickedVertices, _
                                            AdjacentPolys, AdjacentVerts, _
                                            AdjacentAdjacentPolys

                    End If
                Else
                    NumPickedVertices = 0
                End If
            End If
        Case K_MOVE_VERTEX:
            'Freehand vertex movement
            If Ev = K_MOVE Then
                If NumPickedVertices > 0 Then
                    AddStateToBuffer

                    For vi = 0 To NumPickedVertices - 1
                        MoveVertex EditedPModel, PickedVertices(vi), x, y, PickedVertexZ
                    Next vi

                    If glIsEnabled(GL_LIGHTING) = GL_TRUE Then
                        UpdateNormals EditedPModel, PickedVertices, _
                                            AdjacentPolys, AdjacentVerts, _
                                            AdjacentAdjacentPolys
                        'ComputeNormals EditedPModel

                    End If
                End If
            End If
        Case K_NEW_POLY:
            'Create new polygon
            If Ev = K_CLICK Then
                vi = GetClosestVertex(EditedPModel, x, y, repZ + PanZ + DIST)

                If vi > -1 Then
                    VTempNewPoly(VCountNewPoly) = vi
                    VCountNewPoly = VCountNewPoly + 1
                End If

                If VCountNewPoly = 3 Then
                    AddStateToBuffer

                    OrderVertices EditedPModel, VTempNewPoly

                    AddPolygon EditedPModel, VTempNewPoly
                    If glIsEnabled(GL_LIGHTING) = GL_TRUE Then _
                        ComputeNormals EditedPModel
                    VCountNewPoly = 0

                    If PalletizedCheck.value = 1 Then
                        n_colors = 0
                        fill_color_table EditedPModel, color_table, n_colors, translation_table_vertex, _
                                         translation_table_polys, threshold
                    End If
                End If
            End If
        Case K_ROTATE:
            If Ev = K_MOVE Then
                Beta = (Beta + x - x_last) Mod 360
                alpha = (alpha + y - y_last) Mod 360
            End If
        Case K_ZOOM:
            If Ev = K_MOVE Then
                DIST = DIST + (y - y_last) * ComputeDiameter(EditedPModel.BoundingBox) / 100
            End If
        Case K_PAN:
            If Ev = K_MOVE Then
                SetCameraPModel EditedPModel, 0, 0, DIST, 0, 0, 0, redX, redY, redZ
                With p_temp
                    .x = x
                    .y = y
                    .z = GetDepthZ(p_temp2)
                End With
                p_temp = GetUnProjectedCoords(p_temp)

                With p_temp
                    PanX = PanX + .x
                    PanY = PanY + .y
                    PanZ = PanZ + .z
                End With

                With p_temp
                    .x = x_last
                    .y = y_last
                    .z = GetDepthZ(p_temp2)
                End With
                p_temp = GetUnProjectedCoords(p_temp)

                With p_temp
                    PanX = PanX - .x
                    PanY = PanY - .y
                    PanZ = PanZ - .z
                End With
            End If
    End Select
End Sub
Private Sub DrawPallete(ByVal Ev As Integer)
    Dim i As Integer
    Dim s_rows As Single
    Dim x As Integer
    Dim y As Integer
    Dim col As Long
    Dim Brightness As Double
    Dim x0, y0 As Integer
    Dim NewBrush As LOGBRUSH
    Dim hNewBrush As Long
    Dim hBrush As Long
    Dim oldb As Long
    Dim pen As Long, oldpen As Long

    If PalletizedCheck.value Then
        NewBrush.lbColor = RGB(0, 0, 0)
        NewBrush.lbStyle = 2
        NewBrush.lbHatch = 3
        hNewBrush = CreateBrushIndirect(NewBrush)
        oldb = SelectObject(palDC, hNewBrush)
        DeleteObject oldb
        Rectangle palDC, 0, 0, PalletePicture.ScaleWidth, PalletePicture.ScaleHeight

        s_rows = 2 * PalletePicture.ScaleHeight / max(n_colors, 1)
        x0 = 0
        y0 = 0
        i = 0
        While i < n_colors
            With color_table(i)
                hBrush = CreateSolidBrush(RGB(.r, .g, .B))
                oldb = SelectObject(palDC, hBrush)
                DeleteObject oldb
            End With
            If i = Selected_color Then
                pen = CreatePen(0, 2, RGB(255, 255, 255))
                oldpen = SelectObject(palDC, pen)
                Rectangle palDC, x0, y0, x0 + PalletePicture.ScaleWidth / 2, y0 + s_rows
                SelectObject palDC, oldpen
                DeleteObject pen
            Else
                    Rectangle palDC, x0, y0, x0 + PalletePicture.ScaleWidth / 2, y0 + s_rows
            End If

            x0 = x0 + PalletePicture.ScaleWidth / 2

            i = i + 1

            If i Mod 2 = 0 Then
                y0 = y0 + s_rows
                x0 = 0
            End If
        Wend
        BitBlt PalletePicture.hdc, 0, 0, PalletePicture.ScaleWidth, PalletePicture.ScaleHeight, palDC, 0, 0, SRCCOPY
    Else
        With PalletePicture
            For x = 0 To PalletePicture.ScaleWidth - 1
                If Ev = K_LOAD Then
                    For y = 0 To PalletePicture.ScaleHeight * 0.9
                        SetPixel FullPalDC, x, y, RGB((x / .ScaleWidth) * 255, (y / .ScaleHeight) * 255 / 0.9, _
                                                 ((.ScaleWidth - x) / .ScaleWidth) * 255)
                    Next y
                End If
                BitBlt palDC, 0, 0, .ScaleWidth, .ScaleHeight * 0.9, FullPalDC, 0, 0, SRCCOPY
                Brightness = 2 * (x - .ScaleWidth / 2) / .ScaleWidth + 1
                col = RGB(SelectedColorR.value * Brightness, SelectedColorG.value * Brightness, SelectedColorB.value * Brightness)
                For y = PalletePicture.ScaleHeight * 0.9 To PalletePicture.ScaleHeight - 1
                    SetPixel palDC, x, y, col
                Next y
            Next x
            BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, palDC, 0, 0, SRCCOPY
        End With

    End If
End Sub
Private Sub ResizeX_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    redX = ResizeX.value / 100
    ResizeXText.Text = ResizeX.value
    EditedPModel.ResizeX = redX
    Call MainPicture_Paint
    DoNotAddStateQ = False
End Sub
Private Sub ResizeY_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    redY = ResizeY.value / 100
    ResizeYText.Text = ResizeY.value
    EditedPModel.ResizeY = redY
    Call MainPicture_Paint
    DoNotAddStateQ = False
End Sub
Private Sub ResizeZ_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    redZ = ResizeZ.value / 100
    ResizeZText.Text = ResizeZ.value
    EditedPModel.ResizeZ = redZ
    Call MainPicture_Paint
    DoNotAddStateQ = False
End Sub
Private Sub RepositionZ_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    repZ = RepositionZ.value * EditedPModel.diameter / 100
    RepositionZText.Text = RepositionZ.value
    EditedPModel.RepositionZ = repZ
    Call MainPicture_Paint
    DoNotAddStateQ = False
End Sub
Private Sub RepositionX_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    repX = RepositionX.value * EditedPModel.diameter / 100
    RepositionXText.Text = RepositionX.value
    EditedPModel.RepositionX = repX
    Call MainPicture_Paint
    DoNotAddStateQ = False
End Sub

Private Sub RotateAlpha_Change()
    RotationModifiersChanged
End Sub

Private Sub RotateBeta_Change()
    RotationModifiersChanged
End Sub

Private Sub RotateGamma_Change()
    RotationModifiersChanged
End Sub
Private Sub RepositionY_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    repY = RepositionY.value * EditedPModel.diameter / 100
    RepositionYText.Text = RepositionY.value
    EditedPModel.RepositionY = repY
    Call MainPicture_Paint
    DoNotAddStateQ = False
End Sub
Private Sub RepositionXText_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If IsNumeric(RepositionXText.Text) Then
        If RepositionXText.Text >= -100 And RepositionXText.Text <= 100 Then
            RepositionX.value = RepositionXText.Text
            Call RepositionX_Change
        Else
            RepositionXText.Text = 100
            Call RepositionX_Change
        End If
    Else
        Beep
        RepositionXText.Text = 0
        Call RepositionX_Change
    End If
End Sub

Private Sub RotateAlphaText_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If IsNumeric(RotateAlphaText.Text) Then
        If RotateAlphaText.Text >= 0 And RotateAlphaText.Text <= 360 Then
            RotateAlpha.value = RotateAlphaText.Text
            Call RotateAlpha_Change
        Else
            RotateAlphaText.Text = 360
            Call RotateAlpha_Change
        End If
    Else
        Beep
        RotateAlphaText.Text = 0
        Call RotateAlpha_Change
    End If
End Sub

Private Sub RotateBetaText_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If IsNumeric(RotateBetaText.Text) Then
        If RotateBetaText.Text >= 0 And RotateBetaText.Text <= 360 Then
            RotateBeta.value = RotateBetaText.Text
            Call RotateBeta_Change
        Else
            RotateBetaText.Text = 360
            Call RotateBeta_Change
        End If
    Else
        Beep
        RotateBetaText.Text = 0
        Call RotateBeta_Change
    End If
End Sub


Private Sub RotateGammaText_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If IsNumeric(RotateGammaText.Text) Then
        If RotateGammaText.Text >= 0 And RotateGammaText.Text <= 360 Then
            RotateGamma.value = RotateGammaText.Text
            Call RotateGamma_Change
        Else
            RotateGammaText.Text = 360
            Call RotateGamma_Change
        End If
    Else
        Beep
        RotateGammaText.Text = 0
        Call RotateGamma_Change
    End If
End Sub

Private Sub resizexText_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If IsNumeric(ResizeXText.Text) Then
        If ResizeXText.Text >= 0 And ResizeXText.Text <= 400 Then
            ResizeX.value = ResizeXText.Text
            ResizeX_Change
        Else
            ResizeXText.Text = 400
            ResizeX_Change
        End If
    Else
        Beep
        ResizeXText.Text = 0
        ResizeX_Change
    End If
End Sub
Private Sub resizeyText_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If IsNumeric(ResizeYText.Text) Then
        If ResizeYText.Text >= 0 And ResizeYText.Text <= 400 Then
            ResizeY.value = ResizeYText.Text
            ResizeY_Change
        Else
            ResizeYText.Text = 400
            ResizeY_Change
        End If
    Else
        Beep
        ResizeYText.Text = 0
        ResizeY_Change
    End If
End Sub

Private Sub ResizeZText_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If IsNumeric(ResizeZText.Text) Then
        If ResizeZText.Text >= 0 And ResizeZText.Text <= 400 Then
            ResizeZ.value = ResizeZText.Text
            ResizeZ_Change
        Else
            ResizeZText.Text = 400
            ResizeZ_Change
        End If
    Else
        Beep
        ResizeZText.Text = 0
        ResizeZ_Change
    End If
End Sub

Private Sub RepositionZText_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If IsNumeric(RepositionZText.Text) Then
        If RepositionZText.Text >= -100 And RepositionZText.Text <= 100 Then
            RepositionZ.value = RepositionZText.Text
            Call RepositionZ_Change
        Else
            RepositionZText.Text = 100
            Call RepositionZ_Change
        End If
    Else
        Beep
        RepositionZText.Text = 0
        Call RepositionZ_Change
    End If
End Sub

Private Sub RepositionYText_Change()
    If LoadingModifiersQ Then _
        Exit Sub

    If IsNumeric(RepositionYText.Text) Then
        If RepositionYText.Text >= -100 And RepositionYText.Text <= 100 Then
            RepositionY.value = RepositionYText.Text
            Call RepositionY_Change
        Else
            RepositionYText.Text = 100
            Call RepositionY_Change
        End If
    Else
        Beep
        RepositionYText.Text = 0
        Call RepositionY_Change
    End If
End Sub
Public Sub OpenP(ByRef fileName As String)
    On Error GoTo hand
    Dim models3ds_auxV() As Model3Ds
    Dim temp_p As PModel
    Dim p_min As Point3D
    Dim p_max As Point3D

    MainPicture.Enabled = False
    If fileName <> "" Then
        If LCase(Right$(fileName, 4)) = ".3ds" Then
            Load3DS fileName, models3ds_auxV
            ConvertModels3DsToPModel models3ds_auxV, temp_p
        Else
            ReadPModel temp_p, fileName
        End If
    End If
    If fileName = "" Or temp_p.head.NumVerts > 0 Then
        If temp_p.head.NumVerts > 0 Then EditedPModel = temp_p
        'Invalidate the undo/redo buffer
        UnDoCursor = 0
        ReDoCursor = 0

        ComputeBoundingBox EditedPModel
        'ComputeNormals EditedPModel

        SetOGLContext MainPicture.hdc, OGLContextEditor

        glEnable GL_DEPTH_TEST

        glClearColor 0.5, 0.5, 1, 0

        n_colors = 0
        fill_color_table EditedPModel, color_table, n_colors, translation_table_vertex, translation_table_polys, threshold

        MainPicture_Paint
        DrawPallete K_LOAD

        SelectedColorR.Enabled = True
        SelectedColorG.Enabled = True
        SelectedColorB.Enabled = True

        ResizeZ.Enabled = True
        ResizeY.Enabled = True
        ResizeX.Enabled = True
        RepositionX.Enabled = True
        RepositionY.Enabled = True
        RepositionZ.Enabled = True
        RotateGamma.Enabled = True
        RotateBeta.Enabled = True
        RotateAlpha.Enabled = True

        SelectedColorRText.Enabled = True
        SelectedColorGText.Enabled = True
        SelectedColorBText.Enabled = True
        ThersholdText.Enabled = True

        ResizeXText.Enabled = True
        ResizeYText.Enabled = True
        ResizeZText.Enabled = True

        RepositionZText.Enabled = True
        RepositionYText.Enabled = True
        RepositionXText.Enabled = True

        RotateAlphaText.Enabled = True
        RotateBetaText.Enabled = True
        RotateGammaText.Enabled = True

        redX = 1
        redY = 1
        redZ = 1

        repX = 0
        repY = 0
        repZ = 0


        d_type = 2
        Selected_color = -1
        threshold = 20
        alpha = 0
        Beta = 0
        Gamma = 0
        RotateAlpha.value = 0
        RotateAlphaText.Text = 0
        RotateBeta.value = 0
        RotateBetaText.Text = 0
        RotateGamma.value = 0
        RotateGammaText.Text = 0

        file = EditedPModel.fileName

        yl = 0
        ComputePModelBoundingBox EditedPModel, p_min, p_max
        DIST = -2 * ComputeSceneRadius(p_min, p_max)

        LightXScroll.value = 0

        LightYScroll.value = 0

        LightZScroll.value = 0

        Call MainPicture_Paint
        shifted = False
        loaded = True
        CopyVColors EditedPModel.vcolors, VColors_Original
        LoadingModifiersQ = True
        With EditedPModel
            ResizeZ.value = .ResizeZ * 100
            ResizeY.value = .ResizeY * 100
            ResizeX.value = .ResizeX * 100
            ResizeXText.Text = .ResizeX
            ResizeYText.Text = .ResizeY
            ResizeZText.Text = .ResizeZ
            redX = .ResizeX
            redY = .ResizeY
            redZ = .ResizeZ

            RepositionX.value = .RepositionX / .diameter * 100
            RepositionY.value = .RepositionY / .diameter * 100
            RepositionZ.value = .RepositionZ / .diameter * 100
            RepositionXText.Text = RepositionX.value
            RepositionYText.Text = RepositionY.value
            RepositionZText.Text = RepositionZ.value
            repX = .RepositionX
            repY = .RepositionY
            repZ = .RepositionZ

            RotateGamma.value = .RotateGamma
            RotateBeta.value = .RotateBeta
            RotateAlpha.value = .RotateAlpha
            RotateGammaText.Text = RotateGamma.value
            RotateBetaText.Text = RotateBeta.value
            RotateAlphaText.Text = RotateAlpha.value
        End With
        LoadingModifiersQ = False

        FillGroupsList

        ResetCamera
    End If
    Timer1.Enabled = True

    ResetPlane
    Exit Sub
hand:
    MsgBox "Error openning" + fileName, vbOKOnly, "ERROR!"
    Timer1.Enabled = True
End Sub
Public Sub SaveP(ByRef fileName As String)
    Dim mres As Integer
    Dim p_temp As Point3D
    On Error GoTo hand

    mres = vbYes
    If FileExist(fileName) Then _
        mres = MsgBox("File already exists. Overwrite?" + fileName, vbYesNo, "Overwite prompt")

    If mres = vbYes Then
        ApplyContextualizedPChanges (UCase(Right$(fileName, 2)) <> ".P")

        WritePModel EditedPModel, fileName
    End If
    Exit Sub
hand:
    MsgBox "Error saving" + fileName, vbOKOnly, "ERROR!"
End Sub
Sub ApplyContextualizedPChanges(ByVal DNormals As Boolean)
    Dim p_min As Point3D
    Dim p_max As Point3D
    Dim model_diameter_normalized As Single
    Dim temp_dist As Single

    AddStateToBuffer

    SetCameraModelViewQuat repX, repY, repZ, _
                    EditedPModel.RotationQuaternion, _
                    redX, redY, redZ
    glMatrixMode GL_MODELVIEW
    glPushMatrix

    ComputeCurrentBoundingBox EditedPModel

    ComputePModelBoundingBox EditedPModel, p_min, p_max
    temp_dist = -2 * ComputeSceneRadius(p_min, p_max)

    ComputeNormals EditedPModel

    If LightingCheck.value = 1 Then
        glViewport 0, 0, MainPicture.ScaleWidth, MainPicture.ScaleHeight
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT

        SetCameraAroundModelQuat p_min, p_max, repX, repY, repZ + temp_dist, _
                        EditedPModel.RotationQuaternion, _
                        redX, redY, redZ

        glDisable GL_LIGHT0
        glDisable GL_LIGHT1
        glDisable GL_LIGHT2
        glDisable GL_LIGHT3
        ComputePModelBoundingBox EditedPModel, p_min, p_max
        model_diameter_normalized = (-2 * ComputeSceneRadius(p_min, p_max)) / LIGHT_STEPS
        SetLighting GL_LIGHT0, model_diameter_normalized * LightXScroll.value, _
                                model_diameter_normalized * LightYScroll.value, _
                                model_diameter_normalized * LightZScroll.value, 1, 1, 1, False
        ApplyCurrentVColors EditedPModel
    End If

    glMatrixMode GL_MODELVIEW
    glPopMatrix

    SetCameraModelViewQuat repX, repY, repZ, _
                    EditedPModel.RotationQuaternion, _
                    redX, redY, redZ

    ApplyPChanges EditedPModel, DNormals

    LoadingModifiersQ = True

    ResizeX.value = 100
    ResizeY.value = 100
    ResizeZ.value = 100
    RepositionX.value = 0
    RepositionY.value = 0
    RepositionZ.value = 0
    RotateAlpha.value = 0
    RotateBeta.value = 0
    RotateGamma.value = 0

    RepositionXText.Text = 0
    RepositionYText.Text = 0
    RepositionZText.Text = 0
    ResizeXText.Text = 100
    ResizeYText.Text = 100
    ResizeZText.Text = 100
    RotateAlphaText.Text = 0
    RotateBetaText.Text = 0
    RotateGammaText.Text = 0
    EditedPModel.RotationQuaternion.x = 0
    EditedPModel.RotationQuaternion.y = 0
    EditedPModel.RotationQuaternion.z = 0
    EditedPModel.RotationQuaternion.w = 1

    redX = 1
    redY = 1
    redZ = 1
    repX = 0
    repY = 0
    repZ = 0

    LoadingModifiersQ = False
End Sub
Sub FillGroupsList()
    Dim gi As Integer
    Dim name_str As String

    GroupsList.Clear

    For gi = 0 To EditedPModel.head.NumGroups - 1
        If EditedPModel.Groups(gi).HiddenQ Then
            name_str = "[HIDDEN]"
        Else
            name_str = ""
        End If

        name_str = name_str + "Group " + Str$(gi)

        If EditedPModel.Groups(gi).texFlag = 1 Then _
            name_str = name_str + "(Tex.Index:" + _
                                Str$(EditedPModel.Groups(gi).TexID) + ")"

        GroupsList.AddItem name_str
    Next gi
End Sub
Private Sub AddStateToBuffer()
    Dim si As Integer

    If (UnDoCursor < UNDO_BUFFER_CAPACITY) Then
        StoreState UnDoBuffer(UnDoCursor)
        UnDoCursor = UnDoCursor + 1
    Else
        For si = 0 To UNDO_BUFFER_CAPACITY - 1
            UnDoBuffer(si) = UnDoBuffer(si + 1)
        Next si
        StoreState UnDoBuffer(UnDoCursor)
    End If
    ReDoCursor = 0
End Sub
Private Sub ReDo()
    Dim si As Integer

    If loaded Then
        If ReDoCursor > 0 Then
            If (UnDoCursor < UNDO_BUFFER_CAPACITY) Then
                StoreState UnDoBuffer(UnDoCursor)
                UnDoCursor = UnDoCursor + 1
            Else
                'If we've run out of space delete the oldest iteration
                For si = 0 To UNDO_BUFFER_CAPACITY - 1
                    UnDoBuffer(si) = UnDoBuffer(si + 1)
                Next si
                StoreState UnDoBuffer(UnDoCursor)
            End If

            RestoreState ReDoBuffer(ReDoCursor - 1)
            ReDoCursor = ReDoCursor - 1
        Else
            Beep
        End If
    Else
        Beep
    End If
End Sub
Private Sub UnDo()
    Dim si As Integer

    If loaded Then
        If UnDoCursor > 0 Then
            If (ReDoCursor < UNDO_BUFFER_CAPACITY) Then
                StoreState ReDoBuffer(ReDoCursor)
                ReDoCursor = ReDoCursor + 1
            Else
                'If we've run out of space delete the oldest iteration
                For si = 0 To UNDO_BUFFER_CAPACITY - 1
                    ReDoBuffer(si) = ReDoBuffer(si + 1)
                Next si
                StoreState ReDoBuffer(ReDoCursor)
            End If

            RestoreState UnDoBuffer(UnDoCursor - 1)
            UnDoCursor = UnDoCursor - 1
        Else
            Beep
        End If
    Else
        Beep
    End If
End Sub

Private Sub RestoreState(ByRef State As PEditorState)
    LoadingModifiersQ = True

    With State
        PanX = .PanX
        PanY = .PanY
        PanZ = .PanZ

        DIST = .DIST

        redX = .redX
        redY = .redY
        redZ = .redZ

        repX = .repX
        repY = .repY
        repZ = .repZ

        ResizeX.value = .redX * 100
        ResizeY.value = .redY * 100
        ResizeZ.value = .redZ * 100

        RepositionX.value = .repX / EditedPModel.diameter * 100
        RepositionY.value = .repY / EditedPModel.diameter * 100
        RepositionZ.value = .repZ / EditedPModel.diameter * 100

        RotateAlpha.value = .RotateAlpha
        RotateBeta.value = .RotateBeta
        RotateGamma.value = .RotateGamma

        ResizeXText.Text = ResizeX.value
        ResizeYText.Text = ResizeY.value
        ResizeZText.Text = ResizeZ.value

        RepositionXText.Text = RepositionX.value
        RepositionYText.Text = RepositionY.value
        RepositionZText.Text = RepositionZ.value

        RotateAlphaText.Text = RotateAlpha.value
        RotateBetaText.Text = RotateBeta.value
        RotateGammaText.Text = RotateGamma.value

        alpha = .alpha
        Beta = .Beta
        Gamma = .Gamma

        EditedPModel = .EditedPModel

        If .PalletizedQ Then
            PalletizedCheck.value = vbChecked

            color_table = .color_table
            translation_table_polys = .translation_table_polys
            translation_table_vertex = .translation_table_vertex
            n_colors = .n_colors
            threshold = .threshold
        Else
            PalletizedCheck.value = vbUnchecked
        End If
    End With

    If PalletizedCheck.value = 1 Then
        n_colors = 0
        fill_color_table EditedPModel, color_table, n_colors, translation_table_vertex, _
                         translation_table_polys, threshold
    End If

    GroupProperties.Hide
    FillGroupsList

    LoadingModifiersQ = False
End Sub

Private Sub StoreState(ByRef State As PEditorState)
    With State
        .PanX = PanX
        .PanY = PanY
        .PanZ = PanZ

        .DIST = DIST

        .redX = redX
        .redY = redY
        .redZ = redZ

        .repX = repX
        .repY = repY
        .repZ = repZ

        .RotateAlpha = RotateAlpha.value
        .RotateBeta = RotateBeta.value
        .RotateGamma = RotateGamma.value

        .alpha = alpha
        .Beta = Beta
        .Gamma = Gamma

        .EditedPModel = EditedPModel

        .PalletizedQ = PalletizedCheck.value = vbChecked
        If .PalletizedQ Then
            .color_table = color_table
            .translation_table_polys = translation_table_polys
            .translation_table_vertex = translation_table_vertex
            .n_colors = n_colors
            .threshold = threshold
        End If
    End With
End Sub
Private Sub ResetCamera()
    Dim p_min As Point3D
    Dim p_max As Point3D

    If loaded Then
        ComputePModelBoundingBox EditedPModel, p_min, p_max

        alpha = 0
        Beta = 0
        Gamma = 0
        PanX = 0
        PanY = 0
        PanZ = 0
        DIST = -2 * ComputeSceneRadius(p_min, p_max)
    End If
End Sub

Private Sub DrawAxes()
    Dim letter_width As Single
    Dim letter_height As Single

    Dim p_x As Point3D
    Dim p_y As Point3D
    Dim p_z As Point3D

    Dim p_max As Point3D
    Dim p_min As Point3D

    Dim max_x As Single
    Dim max_y As Single
    Dim max_z As Single

    glDisable GL_LIGHTING
    ComputePModelBoundingBox EditedPModel, p_min, p_max

    max_x = IIf(Abs(p_min.x) > Abs(p_max.x), p_min.x, p_max.x)
    max_y = IIf(Abs(p_min.y) > Abs(p_max.y), p_min.y, p_max.y)
    max_z = IIf(Abs(p_min.z) > Abs(p_max.z), p_min.z, p_max.z)
    glBegin GL_LINES
        glColor3f 1, 0, 0
        glVertex3f 0, 0, 0
        glVertex3f 2 * max_x, 0, 0

        glColor3f 0, 1, 0
        glVertex3f 0, 0, 0
        glVertex3f 0, 2 * max_y, 0

        glColor3f 0, 0, 1
        glVertex3f 0, 0, 0
        glVertex3f 0, 0, 2 * max_z
    glEnd

    'Get projected end of the X axis
    p_x.x = 2 * max_x
    p_x.y = 0
    p_x.z = 0
    p_x = GetProjectedCoords(p_x)

    'Get projected end of the Y axis
    p_y.x = 0
    p_y.y = 2 * max_y
    p_y.z = 0
    p_y = GetProjectedCoords(p_y)

    'Get projected end of the Z axis
    p_z.x = 0
    p_z.y = 0
    p_z.z = 2 * max_z
    p_z = GetProjectedCoords(p_z)

    'Set 2D mode to draw letters
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    gluOrtho2D 0, MainPicture.ScaleWidth, 0, MainPicture.ScaleHeight
    glMatrixMode GL_MODELVIEW
    glLoadIdentity

    letter_width = LETTER_SIZE
    letter_height = LETTER_SIZE * 1.5
    glDisable GL_DEPTH_TEST
    glBegin GL_LINES
        'Draw X
        glColor3f 0, 0, 0
        glVertex2f p_x.x - letter_width, p_x.y - letter_height
        glVertex2f p_x.x + letter_width, p_x.y + letter_height
        glVertex2f p_x.x - letter_width, p_x.y + letter_height
        glVertex2f p_x.x + letter_width, p_x.y - letter_height

        'Draw Y
        glColor3f 0, 0, 0
        glVertex2f p_y.x - letter_width, p_y.y - letter_height
        glVertex2f p_y.x + letter_width, p_y.y + letter_height
        glVertex2f p_y.x - letter_width, p_y.y + letter_height
        glVertex2f p_y.x, p_y.y

        'Draw Z
        glColor3f 0, 0, 0
        glVertex2f p_z.x + letter_width, p_z.y + letter_height
        glVertex2f p_z.x - letter_width, p_z.y + letter_height

        glVertex2f p_z.x + letter_width, p_z.y + letter_height
        glVertex2f p_z.x - letter_width, p_z.y - letter_height

        glVertex2f p_z.x + letter_width, p_z.y - letter_height
        glVertex2f p_z.x - letter_width, p_z.y - letter_height
    glEnd
    glEnable GL_DEPTH_TEST
End Sub
Public Sub SetOGLSettings()
    SetOGLContext MainPicture.hdc, OGLContextEditor

    glClearDepth 1#
    glDepthFunc GL_LEQUAL
    glEnable GL_DEPTH_TEST
    glEnable GL_BLEND
    glEnable GL_ALPHA_TEST
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA

    Call MainPicture_Paint
End Sub

Private Sub ZPlaneText_Change()
   If IsNumeric(ZPlaneText.Text) Then
        If ZPlaneText.Text <= ZPlaneUpDown.max _
        And ZPlaneText.Text >= ZPlaneUpDown.Min Then
            ZPlaneUpDown.value = ZPlaneText.Text
        End If
    Else
        Beep
    End If
End Sub

Private Sub ZPlaneUpDown_Change()
    PlaneTransformation(14) = ZPlaneUpDown.value * EditedPModel.diameter / 100

    ComputeCurrentEquations

    MainPicture_Paint
End Sub

Private Sub DrawPlane()
    Dim p1 As Point3D, p2 As Point3D, p3 As Point3D, p4 As Point3D

    MultiplyPoint3DByOGLMatrix PlaneTransformation, PlaneOriginalPoint1, p1
    MultiplyPoint3DByOGLMatrix PlaneTransformation, PlaneOriginalPoint2, p2
    MultiplyPoint3DByOGLMatrix PlaneTransformation, PlaneOriginalPoint3, p3
    MultiplyPoint3DByOGLMatrix PlaneTransformation, PlaneOriginalPoint4, p4

    glPolygonMode GL_FRONT, GL_FILL
    glPolygonMode GL_BACK, GL_LINE

    glColor4f 1, 0, 0, 0.25
    glBegin GL_QUADS
        With p1
            glVertex3f .x, .y, .z
        End With
        With p2
            glVertex3f .x, .y, .z
        End With
        With p3
            glVertex3f .x, .y, .z
        End With
        With p4
            glVertex3f .x, .y, .z
        End With
    glEnd
End Sub

Private Sub ResetPlane()
    AlphaPlaneText.Text = 0
    BetaPlaneText.Text = 0

    XPlaneText.Text = 0
    YPlaneText.Text = 0
    ZPlaneText.Text = 0

    OldAlphaPlane = 0
    OldBetaPlane = 0
    OldGammaPlane = 0

    PlaneOriginalA = 0
    PlaneOriginalB = 0
    PlaneOriginalC = 1
    PlaneOriginalD = 0

    PlaneA = PlaneOriginalA
    PlaneB = PlaneOriginalB
    PlaneC = PlaneOriginalC
    PlaneD = PlaneOriginalD

    With PlaneOriginalPoint
        .x = 0
        .y = 0
        .z = 0
    End With

    With PlaneRotationQuat
        .x = 0
        .y = 0
        .z = 0
        .w = 1
    End With

    BuildMatrixFromQuaternion PlaneRotationQuat, PlaneTransformation

    With PlaneOriginalPoint1
        .z = 0
        .x = EditedPModel.diameter
        .y = EditedPModel.diameter
    End With

    With PlaneOriginalPoint2
        .z = 0
        .x = -EditedPModel.diameter
        .y = EditedPModel.diameter
    End With

    With PlaneOriginalPoint3
        .z = 0
        .x = -EditedPModel.diameter
        .y = -EditedPModel.diameter
    End With

    With PlaneOriginalPoint4
        .z = 0
        .x = EditedPModel.diameter
        .y = -EditedPModel.diameter
    End With

    ComputeCurrentEquations
End Sub

Private Sub ComputeCurrentEquations()
    Dim normal As Point3D
    Dim normal_aux As Point3D

    With normal_aux
        .x = PlaneOriginalA
        .y = PlaneOriginalB
        .z = PlaneOriginalC
    End With

    With PlanePoint
        .x = PlaneTransformation(12)
        .y = PlaneTransformation(13)
        .z = PlaneTransformation(14)
    End With

    PlaneTransformation(12) = 0
    PlaneTransformation(13) = 0
    PlaneTransformation(14) = 0

    MultiplyPoint3DByOGLMatrix PlaneTransformation, normal_aux, normal
    normal = Normalize(normal)

    With normal
        PlaneA = .x
        PlaneB = .y
        PlaneC = .z
    End With

    With PlanePoint
        PlaneTransformation(12) = .x
        PlaneTransformation(13) = .y
        PlaneTransformation(14) = .z
    End With

    With PlanePoint
        PlaneD = -PlaneA * .x - PlaneB * .y - PlaneC * .z
    End With
End Sub

Private Sub RotationModifiersChanged()
    If LoadingModifiersQ Then _
        Exit Sub

    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    RotatePModelModifiers EditedPModel, RotateAlpha.value, RotateBeta.value, RotateGamma.value
    RotateAlphaText.Text = RotateAlpha.value
    RotateBetaText.Text = RotateBeta.value
    RotateGammaText.Text = RotateGamma.value
    Call MainPicture_Paint
    DoNotAddStateQ = False
End Sub


Private Sub UpGroupCommand_Click()
    Dim groupIndex As Integer
    Dim tmpGroups() As PGroup
    Dim i As Integer
            
    ' Get the group selected in p model
    groupIndex = GroupsList.ListIndex
                                       
    If groupIndex > 0 Then
        ' Prepare the new .p model
        
        tmpGroups = EditedPModel.Groups
        
        ' Here we will reorder the groups of the .p model
        For i = 0 To GroupsList.ListCount - 1
            If i = groupIndex - 1 Then
                tmpGroups(i) = EditedPModel.Groups(groupIndex)
            Else
                If i = groupIndex Then
                    tmpGroups(i) = EditedPModel.Groups(i - 1)
                Else
                    tmpGroups(i) = EditedPModel.Groups(i)
                End If
            End If
            
        Next
        
        ' Now we need to reasing the ordered groups to the edited p model
        EditedPModel.Groups = tmpGroups
        
        ' Refresh Groups List
        GroupProperties.Hide
        FillGroupsList
        
        MainPicture_Paint
        
        ' Let's select in listbox the record.
        GroupsList.ListIndex = groupIndex - 1
    End If
End Sub
    
Private Sub DownGroupCommand_Click()
    Dim groupIndex As Integer
    Dim tmpGroups() As PGroup
    Dim i As Integer
            
    ' Get the group selected in p model
    groupIndex = GroupsList.ListIndex
                                       
    If groupIndex > -1 And groupIndex < GroupsList.ListCount - 1 Then
        ' Prepare the new .p model
        
        tmpGroups = EditedPModel.Groups
        
        ' Here we will reorder the groups of the .p model
        For i = 0 To GroupsList.ListCount - 1
            If i = groupIndex + 1 Then
                tmpGroups(i) = EditedPModel.Groups(groupIndex)
            Else
                If i = groupIndex Then
                    tmpGroups(i) = EditedPModel.Groups(i + 1)
                Else
                    tmpGroups(i) = EditedPModel.Groups(i)
                End If
            End If
            
        Next
        
        ' Now we need to reasing the ordered groups to the edited p model
        EditedPModel.Groups = tmpGroups
        
        ' Refresh Groups List
        GroupProperties.Hide
        FillGroupsList
        
        MainPicture_Paint
        
        ' Let's select in listbox the record
        GroupsList.ListIndex = groupIndex + 1
    End If
End Sub



