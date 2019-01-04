VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ModelEditor 
   Caption         =   "Kimera - FF7PC simple model editor v0.98b by Borde - maintained by RedXIV"
   ClientHeight    =   9190
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   9830
   Icon            =   "Main_observer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   613
   ScaleMode       =   0  'User
   ScaleWidth      =   656.313
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SaveFF7AnimationButton 
      Caption         =   "Save Animation"
      Height          =   255
      Left            =   8040
      TabIndex        =   107
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton InterpolateAllAnimsCommand 
      Caption         =   "Interpolate all anims"
      Height          =   255
      Left            =   8040
      TabIndex        =   106
      Top             =   8880
      Width           =   1695
   End
   Begin VB.CommandButton ShowCharModelDBButton 
      Caption         =   "Show Char.lgp DB"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   105
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton InterpolateAnimationButton 
      Caption         =   "Interpolate Animation"
      Height          =   255
      Left            =   8040
      TabIndex        =   104
      Top             =   8640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton ComputeWeaponPositionButton 
      Caption         =   "Compute attached weapon position"
      Height          =   435
      Left            =   0
      TabIndex        =   98
      Top             =   7800
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton ComputeGroundHeightButton 
      Caption         =   "Compute ground height"
      Height          =   255
      Left            =   0
      TabIndex        =   97
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox ShowLastFrameGhostCheck 
      Caption         =   "Overlap last frame ghost"
      Height          =   255
      Left            =   0
      TabIndex        =   96
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton ChangeAnimationButton 
      Caption         =   "Load field animation"
      Height          =   375
      Left            =   8040
      TabIndex        =   91
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox ShowGroundCheck 
      Caption         =   "Show ground"
      Height          =   255
      Left            =   0
      TabIndex        =   90
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton ShowTextureOptionsButton 
      Caption         =   "Show Texture options"
      Height          =   255
      Left            =   8040
      TabIndex        =   75
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame TexturesFrame 
      Caption         =   " Textures (Part)"
      Height          =   3015
      Left            =   9840
      TabIndex        =   73
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
      Begin VB.ComboBox TextureSelectCombo 
         Height          =   315
         Left            =   360
         TabIndex        =   89
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton RemoveTextureButton 
         Caption         =   "Remove Texture"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton AddTextureButton 
         Caption         =   "Add Texture"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton ChangeTextureButton 
         Caption         =   "Change Texture"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox ZeroAsTransparent 
         Caption         =   "0 as transparent"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1920
         Width           =   1455
      End
      Begin VB.PictureBox TextureViewer 
         Height          =   1335
         Left            =   120
         ScaleHeight     =   130
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   142
         TabIndex        =   74
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.UpDown MoveTexureUpDown 
         Height          =   375
         Left            =   120
         TabIndex        =   88
         Top             =   1560
         Width           =   240
         _ExtentX        =   459
         _ExtentY        =   670
         _Version        =   393216
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
   End
   Begin VB.ComboBox WeaponCombo 
      Height          =   315
      Left            =   1320
      TabIndex        =   70
      Text            =   "0"
      Top             =   8640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox BattleAnimationCombo 
      Height          =   315
      Left            =   1320
      TabIndex        =   67
      Text            =   "0"
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox DListEnableCheck 
      Caption         =   "Render using DLists"
      Height          =   255
      Left            =   0
      TabIndex        =   66
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton AnimationOptionsButton 
      Caption         =   "Show Frame options"
      Height          =   255
      Left            =   8040
      TabIndex        =   65
      Top             =   8400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame AnimationOptionsFrame 
      Caption         =   "Frame options"
      Height          =   2895
      Left            =   9840
      TabIndex        =   57
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
      Begin VB.CommandButton InterpolateFrameButton 
         Caption         =   "Interpolate Frame"
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton DuplicateFrameButton 
         Caption         =   "Duplicate Frame"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton RemoveFrameButton 
         Caption         =   "Remove Frame"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox PropagateChangesForwardCheck 
         Caption         =   "Propagate f."
         Height          =   255
         Left            =   240
         TabIndex        =   93
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin MSComCtl2.UpDown FrameDataPartUpDown 
         Height          =   255
         Left            =   1320
         TabIndex        =   83
         Top             =   240
         Width           =   240
         _ExtentX        =   459
         _ExtentY        =   459
         _Version        =   393216
         Max             =   2
         Wrap            =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.Frame FrameDataPartOptions 
         Caption         =   "Bone Rotation"
         Height          =   1335
         Left            =   120
         TabIndex        =   58
         Top             =   480
         Width           =   1455
         Begin MSComCtl2.UpDown ZAnimationFramePartUpDown 
            Height          =   255
            Left            =   1080
            TabIndex        =   82
            Top             =   960
            Width           =   240
            _ExtentX        =   459
            _ExtentY        =   459
            _Version        =   393216
            Increment       =   10000
            Max             =   7200000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown YAnimationFramePartUpDown 
            Height          =   255
            Left            =   1080
            TabIndex        =   81
            Top             =   600
            Width           =   240
            _ExtentX        =   459
            _ExtentY        =   459
            _Version        =   393216
            Increment       =   10000
            Max             =   7200000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown XAnimationFramePartUpDown 
            Height          =   255
            Left            =   1080
            TabIndex        =   80
            Top             =   240
            Width           =   240
            _ExtentX        =   459
            _ExtentY        =   459
            _Version        =   393216
            Increment       =   10000
            Max             =   7200000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin VB.TextBox YAnimationFramePartText 
            Height          =   285
            Left            =   360
            TabIndex        =   61
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox ZAnimationFramePartText 
            Height          =   285
            Left            =   360
            TabIndex        =   60
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox XAnimationFramePartText 
            Height          =   285
            Left            =   360
            TabIndex        =   59
            Top             =   240
            Width           =   735
         End
         Begin VB.Label ZAnimationFramePartLabel 
            AutoSize        =   -1  'True
            Caption         =   "Z"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   960
            Width           =   105
         End
         Begin VB.Label YAnimationFramePartLabel 
            AutoSize        =   -1  'True
            Caption         =   "Y"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   600
            Width           =   105
         End
         Begin VB.Label XAnimationFramePartLabel 
            AutoSize        =   -1  'True
            Caption         =   "X"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   105
         End
      End
      Begin VB.Label FrameOptionsPart 
         AutoSize        =   -1  'True
         Caption         =   "Frame data part"
         Height          =   195
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.CheckBox ShowBonesCheck 
      Caption         =   "Show Bones"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6480
      TabIndex        =   56
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame SelectedPieceFrame 
      Caption         =   "Selected Piece"
      Enabled         =   0   'False
      Height          =   6495
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   2175
      Begin VB.Frame ResizeFrame 
         Caption         =   "Resize"
         Height          =   2055
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1935
         Begin VB.TextBox ResizePieceZText 
            Height          =   285
            Left            =   1320
            TabIndex        =   52
            Text            =   "100"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox ResizePieceYText 
            Height          =   285
            Left            =   1320
            TabIndex        =   51
            Text            =   "100"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox ResizePieceXText 
            Height          =   285
            Left            =   1320
            TabIndex        =   50
            Text            =   "100"
            Top             =   480
            Width           =   375
         End
         Begin VB.HScrollBar ResizePieceZ 
            Height          =   255
            Left            =   240
            Max             =   400
            TabIndex        =   49
            Top             =   1680
            Value           =   100
            Width           =   975
         End
         Begin VB.HScrollBar ResizePieceY 
            Height          =   255
            Left            =   240
            Max             =   400
            TabIndex        =   48
            Top             =   1080
            Value           =   100
            Width           =   975
         End
         Begin VB.HScrollBar ResizePieceX 
            Height          =   255
            Left            =   240
            Max             =   400
            TabIndex        =   47
            Top             =   480
            Value           =   100
            Width           =   975
         End
         Begin VB.Label ResizePieceZLabel 
            Caption         =   "Z re-size"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label ResizePieceYLabel 
            Caption         =   "Y re-size"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label ResizePieceXLabel 
            Caption         =   "X re-size"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame RotateFrame 
         Caption         =   "Rotation"
         Height          =   2055
         Left            =   120
         TabIndex        =   36
         Top             =   4320
         Width           =   1935
         Begin VB.HScrollBar RotateGamma 
            Height          =   255
            Left            =   240
            Max             =   360
            TabIndex        =   42
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox RotateGammaText 
            Height          =   285
            Left            =   1320
            TabIndex        =   41
            Text            =   "0"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox RotateBetaText 
            Height          =   285
            Left            =   1320
            TabIndex        =   40
            Text            =   "0"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox RotateAlphaText 
            Height          =   285
            Left            =   1320
            TabIndex        =   39
            Text            =   "0"
            Top             =   480
            Width           =   375
         End
         Begin VB.HScrollBar RotateBeta 
            Height          =   255
            Left            =   240
            Max             =   360
            TabIndex        =   38
            Top             =   1080
            Width           =   975
         End
         Begin VB.HScrollBar RotateAlpha 
            Height          =   255
            Left            =   240
            Max             =   360
            TabIndex        =   37
            Top             =   480
            Width           =   975
         End
         Begin VB.Label RotateGammaLabel 
            Caption         =   "Gama rotation (Z-axis)"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label RotateBetaLabel 
            Caption         =   "Beta rotation (Y-axis)"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label RotateAlphaLabel 
            Caption         =   "Alpha rotation (X-axis)"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame RepositionFrame 
         Caption         =   "Reposition"
         Height          =   2055
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   1935
         Begin VB.TextBox RepositionZText 
            Height          =   285
            Left            =   1320
            TabIndex        =   32
            Text            =   "0"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox RepositionYText 
            Height          =   285
            Left            =   1320
            TabIndex        =   31
            Text            =   "0"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox RepositionXText 
            Height          =   285
            Left            =   1320
            TabIndex        =   30
            Text            =   "0"
            Top             =   480
            Width           =   375
         End
         Begin VB.HScrollBar RepositionZ 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   -100
            TabIndex        =   29
            Top             =   1680
            Width           =   975
         End
         Begin VB.HScrollBar RepositionY 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   -100
            TabIndex        =   28
            Top             =   1080
            Width           =   975
         End
         Begin VB.HScrollBar RepositionX 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   -100
            TabIndex        =   27
            Top             =   480
            Width           =   975
         End
         Begin VB.Label RepositionZLabel 
            Caption         =   "Z re-position"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label RepositionYLabel 
            Caption         =   "Y re-position"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label RepositionXLabel 
            Caption         =   "X re-position"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Frame GeneralLightingFrame 
      Caption         =   "General Lighting"
      Height          =   2295
      Left            =   8040
      TabIndex        =   17
      Top             =   5520
      Width           =   1695
      Begin VB.CheckBox InifintyFarLightsCheck 
         Caption         =   "Ininitely far lights"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   1920
         Width           =   1455
      End
      Begin VB.HScrollBar LightPosZScroll 
         Height          =   255
         Left            =   360
         TabIndex        =   101
         Top             =   960
         Width           =   1215
      End
      Begin VB.HScrollBar LightPosYScroll 
         Height          =   255
         Left            =   360
         TabIndex        =   100
         Top             =   600
         Width           =   1215
      End
      Begin VB.HScrollBar LightPosXScroll 
         Height          =   255
         Left            =   360
         TabIndex        =   99
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox RightLight 
         Caption         =   "Right"
         Height          =   195
         Left            =   840
         TabIndex        =   24
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox LeftLight 
         Caption         =   "Left"
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox RearLight 
         Caption         =   "Rear"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox FrontLight 
         Caption         =   "Front"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label LightPosZLabel 
         Caption         =   "Z:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   135
      End
      Begin VB.Label LightPosYLabel 
         Caption         =   "Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   135
      End
      Begin VB.Label LightPosXLabel 
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.ComboBox BoneSelector 
      Height          =   315
      Left            =   8040
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton SaveFF7ModelButton 
      Caption         =   "Save Model"
      Height          =   255
      Left            =   8040
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame SelectedBoneFrame 
      Caption         =   "Bone options"
      Enabled         =   0   'False
      Height          =   3015
      Left            =   8040
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
      Begin MSComCtl2.UpDown BoneLengthUpDown 
         Height          =   255
         Left            =   1320
         TabIndex        =   79
         Top             =   1320
         Width           =   240
         _ExtentX        =   459
         _ExtentY        =   459
         _Version        =   393216
         Max             =   999999999
         Min             =   -999999999
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ResizeBoneZUpDown 
         Height          =   285
         Left            =   1320
         TabIndex        =   78
         Top             =   960
         Width           =   240
         _ExtentX        =   459
         _ExtentY        =   512
         _Version        =   393216
         BuddyControl    =   "ResizeBoneZText"
         BuddyDispid     =   196694
         OrigLeft        =   1320
         OrigTop         =   960
         OrigRight       =   1560
         OrigBottom      =   1215
         Max             =   999999999
         Min             =   -999999999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ResizeBoneYUpDown 
         Height          =   285
         Left            =   1320
         TabIndex        =   77
         Top             =   600
         Width           =   240
         _ExtentX        =   459
         _ExtentY        =   512
         _Version        =   393216
         BuddyControl    =   "ResizeBoneYText"
         BuddyDispid     =   196695
         OrigLeft        =   1335
         OrigTop         =   600
         OrigRight       =   1575
         OrigBottom      =   885
         Max             =   999999999
         Min             =   -999999999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ResizeBoneXUpDown 
         Height          =   285
         Left            =   1335
         TabIndex        =   76
         Top             =   240
         Width           =   240
         _ExtentX        =   459
         _ExtentY        =   512
         _Version        =   393216
         BuddyControl    =   "ResizeBoneXText"
         BuddyDispid     =   196690
         OrigLeft        =   1320
         OrigTop         =   240
         OrigRight       =   1560
         OrigBottom      =   495
         Max             =   999999999
         Min             =   -999999999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox ResizeBoneXText 
         Height          =   285
         Left            =   840
         TabIndex        =   72
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton RemovePieceButton 
         Caption         =   "Remove part from the bone"
         Height          =   495
         Left            =   120
         TabIndex        =   69
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton AddPieceButton 
         Caption         =   "Add part to the bone"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox BoneLengthText 
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox ResizeBoneZText 
         Height          =   285
         Left            =   840
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox ResizeBoneYText 
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.Label BoneLengthLabel 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Z Scale"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Y Scale"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "X Scale"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.TextBox AnimationFrameText 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   8880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.HScrollBar CurrentFrameScroll 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3960
      Max             =   0
      TabIndex        =   2
      Top             =   8880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton OpenFF7ModelButton 
      Caption         =   "Open Model"
      Height          =   255
      Left            =   8040
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   8775
      Left            =   2280
      Negotiate       =   -1  'True
      ScaleHeight     =   874
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   562
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   480
         Top             =   0
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label WeaponLabel 
      Caption         =   "Weapon:"
      Height          =   255
      Left            =   0
      TabIndex        =   71
      Top             =   8640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label BattleAnimationLabel 
      Caption         =   "Battle Animation:"
      Height          =   255
      Left            =   0
      TabIndex        =   68
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label BoneSelectorLabel 
      Caption         =   "Selected bone:"
      Height          =   255
      Left            =   8040
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label AnimationFrameLabel 
      Caption         =   "Animation Frame"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "ModelEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const K_P_FIELD_MODEL = 0
Private Const K_P_BATTLE_MODEL = 1
Private Const K_HRC_SKELETON = 2
Private Const K_AA_SKELETON = 3

Private Const K_FRAME_BONE_ROTATION = 0
Private Const K_FRAME_ROOT_ROTATION = 1
Private Const K_FRAME_ROOT_TRANSLATION = 2

Private Const LIGHT_STEPS = 20

Dim ModelType As Integer

Dim P_Model As PModel

Dim hrc_sk As HRCSkeleton
Dim AAnim As AAnimation

Dim aa_sk As AASkeleton
Dim DAAnims As DAAnimationsPack

Dim CurrentFrame As Integer

Dim SelectedBone As Integer
Dim SelectedBonePiece As Integer

Dim EditedBone As Integer
Dim EditedBonePiece As Integer

Dim diameter As Double

Dim alpha As Double
Dim Beta As Double
Dim Gamma As Double

Dim DIST As Double

Dim PanX As Single
Dim PanY As Single
Dim PanZ As Single

Dim loaded As Boolean
Dim rotate As Boolean
Dim SelectBoneForWeaponAttachmentQ As Boolean

Dim x_last As Integer
Dim y_last As Integer
Dim frame_i As Integer
Dim tex_ids(0) As Long
Dim Editor As PEditor
Private MinFormWidth As Single
Private MinFormHeight As Single

Dim LoadingBoneModifiersQ As Boolean
Dim LoadingBonePieceModifiersQ As Boolean
Dim LoadingAnimationQ As Boolean
Dim DoNotAddStateQ As Boolean

Dim ControlPressedQ As Boolean

Dim DontRefreshPicture As Boolean

Private Type ModelEditorState
    P_Model As PModel
    aa_sk As AASkeleton
    hrc_sk As HRCSkeleton
    
    DAAnims As DAAnimationsPack
    AAnim As AAnimation
    
    FrameIndex As Integer
    BattleAnimationIndex As Integer
    WeaponIndex As Integer
    TextureIndex As Integer

    SelectedBone As Integer
    SelectedBonePiece As Integer
    
    EditedBone As Integer
    EditedBonePiece As Integer
    
    alpha As Double
    Beta As Double
    Gamma As Double
    
    DIST As Double
    
    PanX As Single
    PanY As Single
    PanZ As Single
End Type

Dim UnDoBuffer() As ModelEditorState
Dim ReDoBuffer() As ModelEditorState
    
Dim UnDoCursor As Integer
Dim ReDoCursor As Integer

Public Sub OpenFF7File(ByVal fileName As String)
    On Error GoTo errorH
    Dim p_min As Point3D
    Dim p_max As Point3D

    Dim ai As Integer
    Dim wi As Integer
    Dim num_weapons As Integer
    Dim anim_index As Integer
    
    Dim models3ds_auxV() As Model3Ds
    
    Dim p_model_aux As PModel
    
    P_Model = p_model_aux
    
    DisableOpenGL OGLContext
    OGLContext = CreateOGLContext(Picture1.hdc)
    SetOGLContext Picture1.hdc, OGLContext
    ''Debug.Print "Initializing OGL ...", glGetError = GL_NO_ERROR
    
    glEnable GL_DEPTH_TEST
    
    glClearColor 0.5, 0.5, 1, 0
    
    Picture1.Cls
    Picture1.Print "Loading ", fileName
    
    SelectedBone = -1
    SelectedBonePiece = -1
    
    SelectedBoneFrame.Enabled = False
    SelectedPieceFrame.Enabled = False
    
    If LCase(Right$(fileName, 2)) = ".p" Then
        ReadPModel P_Model, fileName
        If P_Model.head.NumVerts > 0 Then
            UnDoCursor = 0
            ReDoCursor = 0
            
            ComputePModelBoundingBox P_Model, p_min, p_max
            diameter = ComputeDiameter(P_Model.BoundingBox)
            With CurrentFrameScroll
                .value = 0
                .Min = -1
                .max = 0
                .Enabled = False
            End With
            ModelType = K_P_FIELD_MODEL
            AddPieceButton.Enabled = False
            ChangeAnimationButton.Enabled = False
            ShowBonesCheck.Enabled = False
            AnimationOptionsFrame.Left = ModelEditor.ScaleWidth + 10
            
            BattleAnimationCombo.Visible = False
            BattleAnimationLabel.Visible = False
            WeaponCombo.Visible = False
            WeaponLabel.Visible = False
            
            SaveFF7ModelButton.Visible = True
            SaveFF7AnimationButton.Visible = False
            ChangeAnimationButton.Visible = False
            SelectedBoneFrame.Visible = False
            ShowTextureOptionsButton.Visible = False
            AnimationOptionsButton.Visible = False
            BoneSelectorLabel.Visible = False
            BoneSelector.Visible = False
            SelectedBoneFrame.Visible = False
            AnimationFrameLabel.Visible = False
            AnimationFrameText.Visible = False
            CurrentFrameScroll.Visible = False
            ShowBonesCheck.Visible = False
            SelectedPieceFrame.Enabled = True
            ComputeGroundHeightButton.Visible = False
            ComputeWeaponPositionButton.Visible = False
            InterpolateAnimationButton.Visible = False
        End If
    Else
        If LCase(Right$(fileName, 4)) = ".hrc" Then
            UnDoCursor = 0
            ReDoCursor = 0
            
            ReadHRCSkeleton hrc_sk, fileName, True
            ReadFirstCompatibleAAnimation
            ComputeHRCBoundingBox hrc_sk, AAnim.Frames(0), p_min, p_max
            diameter = ComputeHRCDiameter(hrc_sk)
            With CurrentFrameScroll
                .value = 0
                .Min = -1
                .max = AAnim.NumFrames
                .Enabled = True
            End With
            ModelType = K_HRC_SKELETON
            AddPieceButton.Enabled = True
            ChangeAnimationButton.Enabled = True
            ShowBonesCheck.Enabled = True
            
            BattleAnimationCombo.Visible = False
            BattleAnimationLabel.Visible = False
            WeaponCombo.Visible = False
            WeaponLabel.Visible = False
            
            SaveFF7ModelButton.Visible = True
            SaveFF7AnimationButton.Visible = True
            SaveFF7AnimationButton.Caption = "Save animation"
            AnimationOptionsFrame.Visible = True
            TextureSelectCombo.Clear
            TexturesFrame.Visible = True
            TexturesFrame.Caption = "Textures (part)"
            ChangeAnimationButton.Visible = True
            ChangeAnimationButton.Caption = "Load field animation"
            ShowTextureOptionsButton.Visible = True
            AnimationOptionsButton.Visible = True
            BoneSelectorLabel.Visible = True
            BoneSelector.Visible = True
            SelectedBoneFrame.Visible = True
            AnimationFrameLabel.Visible = True
            AnimationFrameText.Visible = True
            CurrentFrameScroll.Visible = True
            ShowBonesCheck.Visible = True
            ComputeGroundHeightButton.Visible = True
            ComputeWeaponPositionButton.Visible = False
            InterpolateAnimationButton.Visible = True
            Call SetFrameEditorFields
            Call SetTextureEditorFields
        Else
            If LCase(Right$(fileName, 2)) = "aa" Then
                UnDoCursor = 0
                ReDoCursor = 0
                
                ReadAASkeleton fileName, aa_sk, False, True
                num_weapons = aa_sk.NumWeapons
                WeaponCombo.Clear
                For wi = 0 To num_weapons - 1
                    WeaponCombo.AddItem (wi)
                Next wi
                If aa_sk.IsBattleLocation Then
                    CreateEmptyDAAnimationsPack DAAnims, aa_sk.NumBones + 1
                Else
                    If FileExist(LCase(Left$(Right$(fileName, 4), 2)) + "da") Then
                        ReadDAAnimationsPack LCase(Left$(Right$(fileName, 4), 2)) + "da", aa_sk.NumBones, aa_sk.NumBodyAnims, aa_sk.NumWeaponAnims, DAAnims
                    Else
                        CreateCompatibleDAAnimationsPack aa_sk, DAAnims
                    End If
                End If
                ModelType = K_AA_SKELETON
                ShowBonesCheck.Enabled = True
                BattleAnimationCombo.Clear
                For ai = 0 To DAAnims.NumBodyAnimations - 1
                    If DAAnims.BodyAnimations(ai).NumFrames2 > 0 Then _
                        BattleAnimationCombo.AddItem (ai)
                Next ai
                
                If Not aa_sk.IsBattleLocation Then _
                    anim_index = val(BattleAnimationCombo.List(0))
                ComputeAABoundingBox aa_sk, _
                    DAAnims.BodyAnimations(anim_index).Frames(0), _
                    p_min, p_max
                diameter = ComputeAADiameter(aa_sk)
                    
                SelectedBone = -1
                SelectedBonePiece = -1
                If aa_sk.IsBattleLocation Then
                    BattleAnimationCombo.Visible = False
                    BattleAnimationLabel.Visible = False
                Else
                    BattleAnimationCombo.ListIndex = 0
                    BattleAnimationCombo.Visible = True
                    BattleAnimationLabel.Visible = True
                End If
                
                With CurrentFrameScroll
                    .value = 0
                    .Min = -1
                    .max = max(1, DAAnims.BodyAnimations(0).NumFrames2)
                    .Enabled = True
                End With
                
                AddPieceButton.Enabled = True
                If num_weapons > 0 Then
                    WeaponCombo.ListIndex = 0
                    WeaponCombo.Visible = True
                    WeaponLabel.Visible = True
                    WeaponCombo.ListIndex = 0
                Else
                    WeaponCombo.Visible = False
                    WeaponLabel.Visible = False
                End If
                
                TexturesFrame.Visible = True
                TexturesFrame.Enabled = True
                TexturesFrame.Caption = "Textures (model)"
                
                SaveFF7ModelButton.Visible = True
                SaveFF7AnimationButton.Visible = True
                SaveFF7AnimationButton.Caption = "Save anims. pack"
                ChangeAnimationButton.Visible = ModelHasLimitBreaks(fileName)
                ChangeAnimationButton.Caption = "Load anims. pack"
                AnimationOptionsFrame.Visible = True
                ShowTextureOptionsButton.Visible = True
                AnimationOptionsButton.Visible = True
                BoneSelectorLabel.Visible = True
                BoneSelector.Visible = True
                SelectedBoneFrame.Visible = True
                AnimationFrameLabel.Visible = True
                AnimationFrameText.Visible = True
                CurrentFrameScroll.Visible = True
                ShowBonesCheck.Visible = True
                ComputeGroundHeightButton.Visible = True
                ComputeWeaponPositionButton.Visible = num_weapons > 0
                InterpolateAnimationButton.Visible = True
                Call SetFrameEditorFields
                Call SetTextureEditorFields
            Else
                If LCase(Right$(fileName, 2)) = ".d" Then
                    UnDoCursor = 0
                    ReDoCursor = 0
                    
                    ReadMagicSkeleton fileName, aa_sk, True
                    
                    If FileExist(LCase(Left$(fileName, Len(fileName) - 1)) + "a00") Then
                        ReadDAAnimationsPack LCase(Left$(fileName, Len(fileName) - 1)) + "a00", aa_sk.NumBones, aa_sk.NumBodyAnims, aa_sk.NumWeaponAnims, DAAnims
                        num_weapons = 0
                    Else
                        CreateCompatibleDAAnimationsPack aa_sk, DAAnims
                    End If
                    
                    ModelType = K_AA_SKELETON
                    ShowBonesCheck.Enabled = True
                    BattleAnimationCombo.Clear
                    For ai = 0 To DAAnims.NumBodyAnimations - 1
                        If DAAnims.BodyAnimations(ai).NumFrames2 > 0 Then _
                            BattleAnimationCombo.AddItem (ai)
                    Next ai
                    BattleAnimationCombo.ListIndex = 0
                    
                    anim_index = BattleAnimationCombo.List(0)
                    ComputeAABoundingBox aa_sk, _
                        DAAnims.BodyAnimations(anim_index).Frames(0), _
                        p_min, p_max
                    diameter = ComputeAADiameter(aa_sk)
                    
                    BattleAnimationCombo.Visible = True
                    BattleAnimationLabel.Visible = True
                    WeaponCombo.Visible = False
                    WeaponLabel.Visible = False
                    
                    AddPieceButton.Enabled = True
                    
                    With CurrentFrameScroll
                        .value = 0
                        .Min = -1
                        .max = DAAnims.BodyAnimations(0).NumFrames2 - 1
                        .Enabled = True
                    End With
                    
                    If num_weapons > 0 Then
                        WeaponCombo.ListIndex = 0
                        WeaponCombo.Visible = True
                        WeaponLabel.Visible = True
                        WeaponCombo.ListIndex = 0
                    Else
                        WeaponCombo.Visible = False
                        WeaponLabel.Visible = False
                    End If
                    
                    SaveFF7ModelButton.Visible = True
                    SaveFF7AnimationButton.Visible = True
                    SaveFF7AnimationButton.Caption = "Save anims. pack"
                    AnimationOptionsFrame.Visible = True
                    TexturesFrame.Visible = True
                    TexturesFrame.Enabled = True
                    TexturesFrame.Caption = "Textures (model)"
                    ChangeAnimationButton.Visible = False
                    ShowTextureOptionsButton.Visible = True
                    AnimationOptionsButton.Visible = True
                    BoneSelectorLabel.Visible = True
                    BoneSelector.Visible = True
                    SelectedBoneFrame.Visible = True
                    AnimationFrameLabel.Visible = True
                    AnimationFrameText.Visible = True
                    CurrentFrameScroll.Visible = True
                    ShowBonesCheck.Visible = True
                    ComputeGroundHeightButton.Visible = True
                    ComputeWeaponPositionButton.Visible = False
                    InterpolateAnimationButton.Visible = True
                    Call SetFrameEditorFields
                    Call SetTextureEditorFields
                Else
                    If LCase(Right$(fileName, 4)) = ".3ds" Then
                        Load3DS fileName, models3ds_auxV
                        ConvertModels3DsToPModel models3ds_auxV, P_Model
                    Else
                        ReadPModel P_Model, fileName
                    End If
                    
                    If P_Model.head.NumVerts > 0 Then
                        UnDoCursor = 0
                        ReDoCursor = 0
                        
                        ComputePModelBoundingBox P_Model, p_min, p_max
                        diameter = ComputeDiameter(P_Model.BoundingBox)
                        ModelType = K_P_BATTLE_MODEL
                        ShowBonesCheck.Enabled = False
                        AddPieceButton.Enabled = False
                        
                        BattleAnimationCombo.Visible = False
                        BattleAnimationLabel.Visible = False
                        WeaponCombo.Visible = False
                        WeaponLabel.Visible = False
                
                        TextureSelectCombo.Clear
                        TexturesFrame.Enabled = False
                        SaveFF7ModelButton.Visible = True
                        SaveFF7AnimationButton.Visible = False
                        ChangeAnimationButton.Visible = False
                        BoneSelectorLabel.Visible = False
                        BoneSelector.Visible = False
                        SelectedBoneFrame.Visible = False
                        AnimationOptionsFrame.Visible = False
                        TexturesFrame.Visible = False
                        ShowTextureOptionsButton.Visible = False
                        AnimationOptionsButton.Visible = False
                        SelectedBoneFrame.Visible = False
                        AnimationFrameLabel.Visible = False
                        AnimationFrameText.Visible = False
                        CurrentFrameScroll.Visible = False
                        ShowBonesCheck.Visible = False
                        SelectedPieceFrame.Enabled = True
                        ComputeGroundHeightButton.Visible = False
                        ComputeWeaponPositionButton.Visible = False
                        InterpolateAnimationButton.Visible = False
                    End If
                End If
            End If
        End If
    End If
    loaded = True
    
    FillBoneSelector
    
    alpha = 200
    Beta = 45
    Gamma = 0
    PanX = 0
    PanY = 0
    PanZ = 0
    DIST = -2 * ComputeSceneRadius(p_min, p_max)
    SelectBoneForWeaponAttachmentQ = False
    
    glClearDepth 1#
    glDepthFunc GL_LEQUAL
    glEnable GL_DEPTH_TEST
    glEnable GL_BLEND
    glEnable GL_ALPHA_TEST
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
    glAlphaFunc GL_GREATER, 0
    glCullFace GL_FRONT
    glEnable GL_CULL_FACE
    
    Picture1_Paint
    Exit Sub
errorH:
    MsgBox "Error while loading file ", vbCritical, "ERROR!!!!"
End Sub
Sub SetFieldModelAnimation(ByRef fileName As String)
    Dim AAnim_tmp As AAnimation
    
    If (fileName <> "") Then
        ReadAAnimation AAnim_tmp, fileName
        FixAAnimation hrc_sk, AAnim_tmp
        If AAnim_tmp.NumBones <> hrc_sk.NumBones Then
            MsgBox "The animation has a wrong number of bones", vbOKOnly, "Error"
            Exit Sub
        Else
            AAnim = AAnim_tmp
            Call SetFrameEditorFields
        End If
        
        CurrentFrameScroll.value = 0
    
        CurrentFrameScroll.max = AAnim.NumFrames
    End If
    
    Picture1_Paint
End Sub

Sub SetBattleModelAnimationsPack(ByRef fileName As String)
    Dim DAAnims_tmp As DAAnimationsPack
    Dim ai As Integer
    
    If (fileName <> "") Then
        If Right$(fileName, 2) = "da" Then
            aa_sk.IsLimitBreak = False
            ReadDAAnimationsPack fileName, aa_sk.NumBones, aa_sk.NumBodyAnims, aa_sk.NumWeaponAnims, DAAnims_tmp
        Else
            aa_sk.IsLimitBreak = True
            ReadDAAnimationsPack fileName, aa_sk.NumBones, 8, 8, DAAnims_tmp
        End If

        DAAnims = DAAnims_tmp
        Call SetFrameEditorFields
        
        BattleAnimationCombo.Clear
        For ai = 0 To DAAnims.NumBodyAnimations - 1
            If DAAnims.BodyAnimations(ai).NumFrames2 > 0 Then _
                BattleAnimationCombo.AddItem (ai)
        Next ai
        SelectedBone = -1
        SelectedBonePiece = -1
        If aa_sk.IsBattleLocation Then
            BattleAnimationCombo.Visible = False
            BattleAnimationLabel.Visible = False
        Else
            BattleAnimationCombo.ListIndex = 0
            BattleAnimationCombo.Visible = True
            BattleAnimationLabel.Visible = True
        End If
        
        With CurrentFrameScroll
            .value = 0
            .Min = -1
            .max = max(1, DAAnims.BodyAnimations(0).NumFrames2)
            .Enabled = True
        End With
    End If
    
    Picture1_Paint
End Sub
Sub ReadFirstCompatibleAAnimation()
    Dim animFile As String
    Dim hrc_file As String
    Dim num_bones As Integer
    Dim anim_foundQ As Boolean
    
    If hrc_sk.NumBones < 2 Then
        CreateCompatibleHRCAAnimation hrc_sk, AAnim
        Exit Sub
    End If
    
    On Error GoTo errorH
    
    hrc_file = LCase(hrc_sk.fileName)
    animFile = LCase(Dir("*.a"))
    anim_foundQ = False
    While Not anim_foundQ And animFile <> ""
        anim_foundQ = IsLexicographicallyGreater(animFile, hrc_file)
        If Not anim_foundQ Then _
            animFile = LCase(Dir())
    Wend

    num_bones = -1
    If anim_foundQ Then
        num_bones = ReadAAnimationNBones(animFile)
    End If
        
    If num_bones <> hrc_sk.NumBones Then
        animFile = Dir("*.a")
        
        num_bones = ReadAAnimationNBones(animFile)
        While num_bones <> hrc_sk.NumBones And animFile <> ""
            animFile = Dir()
            num_bones = ReadAAnimationNBones(animFile)
        Wend
    End If
    
    If num_bones <> hrc_sk.NumBones Then
        MsgBox "There is no animation file that fits the model in the same folder", vbCritical, "Error loading animation"
        CreateCompatibleHRCAAnimation hrc_sk, AAnim
    Else
        ReadAAnimation AAnim, animFile
        FixAAnimation hrc_sk, AAnim
    End If
    Exit Sub
errorH:
    MsgBox "Error while loading field animation ", vbCritical, "ERROR!!!!"
End Sub


Private Sub AddPieceButton_Click()
    Dim models3ds_auxV() As Model3Ds
    Dim AdditionalP As PModel
    Dim pattern As String
    
    On Error GoTo hand
    
    Picture1.Enabled = False
    
    If ModelType = K_HRC_SKELETON Or ModelType = K_AA_SKELETON Then
        pattern = "FF7 Field Part file|*.p|FF7 Battle Part file|*.*|3D Studio model|*.3ds"
        CommonDialog1.Filter = pattern
        CommonDialog1.ShowOpen 'Display the Open File Common Dialog
        If (CommonDialog1.fileName = "") Then Exit Sub
        
        If LCase(Right$(CommonDialog1.fileName, 4)) = ".3ds" Then
            Load3DS CommonDialog1.fileName, models3ds_auxV
            ConvertModels3DsToPModel models3ds_auxV, AdditionalP
        Else
            ReadPModel AdditionalP, CommonDialog1.fileName
        End If
        
        If AdditionalP.head.NumVerts > 0 Then
            AddStateToBuffer
            If ModelType = K_HRC_SKELETON Then
                AddHRCBonePiece hrc_sk.Bones(SelectedBone), AdditionalP
            Else
                AddAABoneModel aa_sk.Bones(SelectedBone), AdditionalP
            End If
        End If
    End If
    Picture1_Paint
    Timer1.Enabled = True
    Exit Sub
hand:
    If Err <> 32755 Then
        MsgBox "Error" + Str$(Err), vbOKOnly, "Unknow error Loading"
    End If
    Timer1.Enabled = True
End Sub

Private Sub AddTextureButton_Click()
    Dim pattern As String
    Dim tex As TEXTexture
    
    On Error GoTo hand
    
    Picture1.Enabled = False
    
    pattern = "Any Image file(*.bmp;*.jpg;*.gif;*.ico;*.rle;*.wmf;*.emf)|*.bmp;*.jpg;*.gif;*.png;*.ico;*.rle;*.wmf;*.emf|TEX texture|*.tex"
    CommonDialog1.Filter = pattern
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen 'Display the Open File Common Dialog

    If (CommonDialog1.fileName <> "") Then
        LoadImageAsTexTexture CommonDialog1.fileName, tex
        Select Case ModelType
            Case K_HRC_SKELETON:
                If SelectedBone > -1 And SelectedBonePiece > -1 Then
                    AddStateToBuffer
                    
                    With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece)
                        ReDim Preserve .textures(IIf(.NumTextures = 0, 0, .NumTextures + 1))
                        .textures(.NumTextures) = tex
                        .NumTextures = .NumTextures + 1
                        SetTextureEditorFields
                        TextureSelectCombo.ListIndex = .NumTextures - 1
                    End With
                End If
            Case K_AA_SKELETON:
                With aa_sk
                    If .NumTextures <= 10 Then
                        AddStateToBuffer
                        
                        ReDim Preserve .textures(IIf(.NumTextures = 0, 0, .NumTextures + 1))
                        ReDim Preserve .TexIDS(IIf(.NumTextures = 0, 0, .NumTextures + 1))
                        .NumTextures = .NumTextures + 1
                        .textures(.NumTextures - 1) = tex
                        .textures(.NumTextures - 1).tex_file = _
                            GetBattleModelTextureFilename(aa_sk, .NumTextures - 1)
                        .TexIDS(.NumTextures - 1) = tex.tex_id
                        SetTextureEditorFields
                        TextureSelectCombo.ListIndex = .NumTextures - 1
                    Else
                        MsgBox "The maximum number of textures for battle models is 10", vbOKOnly, "Error"
                    End If
                End With
        End Select
        Picture1_Paint
    End If
    Timer1.Enabled = True
    Exit Sub
hand:
    If Err <> 32755 Then
        MsgBox "Error" + Str$(Err), vbOKOnly, "Unknow error Loading"
    End If
    Timer1.Enabled = True
End Sub

Private Sub AnimationOptionsButton_Click()
    If AnimationOptionsButton.Caption = "Show Light options" Then
        AnimationOptionsFrame.Left = ModelEditor.ScaleWidth + 10
        GeneralLightingFrame.Left = ModelEditor.ScaleWidth - 115
        AnimationOptionsButton.Caption = "Show Frame options"
    Else
        AnimationOptionsFrame.Left = ModelEditor.ScaleWidth - 115
        GeneralLightingFrame.Left = ModelEditor.ScaleWidth + 10
        AnimationOptionsButton.Caption = "Show Light options"
    End If
End Sub

Private Sub BattleAnimationCombo_Click()
    DontRefreshPicture = True
    CurrentFrameScroll.value = 0
    CurrentFrameScroll.max = DAAnims.BodyAnimations(BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)).NumFrames2
    DontRefreshPicture = False
    Picture1_Paint
End Sub


Private Sub BoneLengthText_Change()
    If LoadingBoneModifiersQ Then _
        Exit Sub
        
    If IsNumeric(BoneLengthText.Text) Then
        BoneLengthUpDown.value = BoneLengthText.Text * 10000
        'Call ResizeBoneZUpDown_Change
    Else
        Beep
    End If

End Sub

Private Sub ChangeTextureButton_Click()
    Dim pattern As String
    Dim tex As TEXTexture
    Dim tex_index As Integer
    
    On Error GoTo hand
    
    Picture1.Enabled = False
    
    pattern = "Any Image file(*.bmp;*.jpg;*.gif;*.ico;*.rle;*.wmf;*.emf)|*.bmp;*.jpg;*.gif;*.png;*.ico;*.rle;*.wmf;*.emf|TEX texture|*.tex"
    CommonDialog1.Filter = pattern
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen 'Display the Open File Common Dialog

    If (CommonDialog1.fileName <> "") Then
        LoadImageAsTexTexture CommonDialog1.fileName, tex
        Select Case ModelType
            Case K_HRC_SKELETON:
                If SelectedBone > -1 Then
                    tex_index = TextureSelectCombo.ListIndex
                    If (tex_index > -1) Then
                        AddStateToBuffer
                        
                        With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece)
                            'This is dirty, but will prevent problems with the undo/redo
                            'UnloadTexture .textures(tex_index)
                            .textures(tex_index) = tex
                            SetTextureEditorFields
                            TextureSelectCombo.ListIndex = tex_index
                        End With
                    Else
                        Beep
                    End If
                End If
            Case K_AA_SKELETON:
                With aa_sk
                    tex_index = TextureSelectCombo.ListIndex
                    If (tex_index > -1) Then
                        AddStateToBuffer
                        
                        'This is dirty, but will prevent problems with the undo/redo
                        'UnloadTexture .textures(tex_index)
                        .textures(tex_index) = tex
                        .textures(tex_index).tex_file = _
                            GetBattleModelTextureFilename(aa_sk, tex_index)
                        .TexIDS(tex_index) = tex.tex_id
                        SetTextureEditorFields
                        TextureSelectCombo.ListIndex = tex_index
                    Else
                        Beep
                    End If
                End With
        End Select
        Picture1_Paint
    End If
    Timer1.Enabled = True
    Exit Sub
hand:
    If Err <> 32755 Then
        MsgBox "Error" + Str$(Err), vbOKOnly, "Unknow error Loading"
    End If
    Timer1.Enabled = True
End Sub


Private Sub ComputeGroundHeightButton_Click()
    Dim p_min As Point3D
    Dim p_max As Point3D
    Dim anim_index As Integer
    Dim fi As Integer
    Dim max_diff As Single
    
    AddStateToBuffer
    max_diff = INFINITY_SINGLE
    Select Case ModelType
        Case K_HRC_SKELETON:
            With AAnim
                For fi = 0 To .NumFrames - 1
                    ComputeHRCBoundingBox hrc_sk, AAnim.Frames(fi), p_min, p_max
                    If max_diff > p_max.y Then max_diff = p_max.y
                Next fi
                
                For fi = 0 To .NumFrames - 1
                    .Frames(fi).RootTranslationY = .Frames(fi).RootTranslationY + max_diff
                Next fi
            End With
        Case K_AA_SKELETON:
            anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
            With DAAnims.BodyAnimations(anim_index)
                For fi = 0 To .NumFrames2 - 1
                    ComputeAABoundingBox aa_sk, _
                        DAAnims.BodyAnimations(anim_index).Frames(fi), _
                        p_min, p_max
                    If max_diff > p_max.y Then max_diff = p_max.y
                Next fi
                
                If max_diff <> 0 Then
                    For fi = 0 To .NumFrames2 - 1
                        .Frames(fi).Y_start = .Frames(fi).Y_start - max_diff
                    Next fi
                        
                    'Also don't forget the weapon frames if available
                    If anim_index < DAAnims.NumWeaponAnimations And aa_sk.NumWeapons > 0 Then
                        For fi = 0 To .NumFrames2 - 1
                            DAAnims.WeaponAnimations(anim_index).Frames(fi).Y_start = _
                                DAAnims.WeaponAnimations(anim_index).Frames(fi).Y_start - max_diff
                        Next fi
                    End If
                End If
            End With
    End Select

    Picture1_Paint
    SetFrameEditorFields
End Sub

Private Sub ComputeWeaponPositionButton_Click()
    SelectBoneForWeaponAttachmentQ = True
    Picture1_Paint
    MsgBox "Please, click (right-click = end, left-click = middle) on the bone you want the weapon to be attached to. Press ESC to cancel.", vbOKOnly, "Select bone"
End Sub

Private Sub DListEnableCheck_Click()
    Picture1_Paint
End Sub


Private Sub DuplicateFrameButton_Click()
    Dim anim_index As Integer
    Dim aframe_aux As AFrame
    Dim daframe_aux As DAFrame
    Dim daframe_w_aux As DAFrame
    Dim fi As Integer
    Dim primary_secondary_counters_coef As Single
    
    AddStateToBuffer
    Select Case ModelType
        Case K_HRC_SKELETON:
            With AAnim
                .NumFrames = .NumFrames + 1
                ReDim Preserve .Frames(.NumFrames - 1)
                For fi = .NumFrames - 2 To CurrentFrameScroll.value Step -1
                    .Frames(fi + 1) = .Frames(fi)
                Next fi
                
                CopyAFrame .Frames(CurrentFrameScroll.value), aframe_aux
                .Frames(CurrentFrameScroll.value) = aframe_aux
                CurrentFrameScroll.max = CurrentFrameScroll.max + 1
            End With
        Case K_AA_SKELETON:
            anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
            With DAAnims.BodyAnimations(anim_index)
                'Numframes1 and NumFrames2 are usually different. Don't know if this is relevant at all, but keep the balance between them just in case
                primary_secondary_counters_coef = .NumFrames1 / .NumFrames2
                .NumFrames2 = .NumFrames2 + 1
                .NumFrames1 = .NumFrames2 * primary_secondary_counters_coef
                ReDim Preserve .Frames(.NumFrames2 - 1)
                
                For fi = .NumFrames2 - 2 To CurrentFrameScroll.value Step -1
                    .Frames(fi + 1) = .Frames(fi)
                Next fi
                
                CopyDAFrame .Frames(CurrentFrameScroll.value), daframe_aux
                .Frames(CurrentFrameScroll.value) = daframe_aux
                
                'Also don't forget about the weapon frames (where available)
                If anim_index < DAAnims.NumWeaponAnimations And aa_sk.NumWeapons > 0 Then
                    ReDim Preserve DAAnims.WeaponAnimations(anim_index).Frames(.NumFrames2 - 1)
                    For fi = .NumFrames2 - 2 To CurrentFrameScroll.value Step -1
                        DAAnims.WeaponAnimations(anim_index).Frames(fi + 1) = _
                            DAAnims.WeaponAnimations(anim_index).Frames(fi)
                    Next fi
                    CopyDAFrame DAAnims.WeaponAnimations(anim_index).Frames(CurrentFrameScroll.value), daframe_w_aux
                    DAAnims.WeaponAnimations(anim_index).Frames(CurrentFrameScroll.value) = daframe_w_aux
                    DAAnims.WeaponAnimations(anim_index).NumFrames2 = .NumFrames2
                    DAAnims.WeaponAnimations(anim_index).NumFrames1 = .NumFrames1
                End If
                    
                CurrentFrameScroll.max = CurrentFrameScroll.max + 1
            End With
    End Select

    Picture1_Paint
End Sub

Private Sub Form_Activate()
    SetOGLSettings
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_HOME Then _
        ResetCamera
        
    If KeyCode = KEY_ESCAPE Then _
        SelectBoneForWeaponAttachmentQ = False
    
    If KeyCode = vbKeyControl Then _
        ControlPressedQ = True
        
    If KeyCode = vbKeyZ And ControlPressedQ Then _
        UnDo
    If KeyCode = vbKeyY And ControlPressedQ Then _
        ReDo
        
    If KeyCode = vbKeyUp Then _
        alpha = alpha + 1
    
    If KeyCode = vbKeyDown Then _
        alpha = alpha - 1
        
    If KeyCode = vbKeyLeft Then _
        Beta = Beta - 1
    
    If KeyCode = vbKeyRight Then _
        Beta = Beta + 1
    
    'Debug.Print alpha; "/"; Beta
        
    Picture1_Paint
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyControl Then _
        ControlPressedQ = False
End Sub

Private Sub FrameDataPartUpDown_Change()
    Select Case FrameDataPartUpDown.value
        Case K_FRAME_BONE_ROTATION
            FrameDataPartOptions.Caption = "Bone rotation"
        Case K_FRAME_ROOT_ROTATION
            FrameDataPartOptions.Caption = "Root rotation"
        Case K_FRAME_ROOT_TRANSLATION
            FrameDataPartOptions.Caption = "Root translation"
    End Select
    SetFrameEditorFields
End Sub

Private Sub InifintyFarLightsCheck_Click()
    Call Picture1_Paint
End Sub

Private Sub InterpolateAllAnimsCommand_Click()
    InterpolateAllAnimsForm.ResetForm
    InterpolateAllAnimsForm.Show vbModal
End Sub

Private Sub InterpolateAnimationButton_Click()
    Dim anim_index As Integer
    Dim aframe_aux As AFrame
    Dim daframe_aux As DAFrame
    Dim daframe_w_aux As DAFrame
    Dim fi As Integer
    Dim ifi As Integer
    Dim prev_frame As Integer
    Dim next_frame As Integer
    Dim is_loop As VbMsgBoxResult
    Dim frame_offset As Integer
    Dim num_frames As Integer
    Dim num_interpolated_frames_str As String
    Dim num_interpolated_frames As Integer
    Dim next_elem_diff As Integer
    Dim alpha As Single
    Dim base_final_frame As Integer
    Dim primary_secondary_counters_coef As Single
    Dim aux_daframe As DAFrame
    
    num_interpolated_frames_str = InputBox("Number of frames to interpolate between each frame", _
                                            "Animation interpolation", _
                                            IIf(ModelType = K_HRC_SKELETON, Str$(DEFAULT_FIELD_INTERP_FRAMES), _
                                                                            Str$(DEFAULT_BATTLE_INTERP_FRAMES)))
    
    If num_interpolated_frames_str = "" Or Not IsNumeric(num_interpolated_frames_str) Then
        Exit Sub
    End If
    
    num_interpolated_frames = CInt(num_interpolated_frames_str)
    next_elem_diff = num_interpolated_frames + 1
    
    AddStateToBuffer
    
    is_loop = MsgBox("Is this animation a loop?", vbYesNo, "Animation type")
    frame_offset = 0
    If is_loop = vbNo Then
        frame_offset = num_interpolated_frames
    End If
    
    Select Case ModelType
        Case K_HRC_SKELETON:
            With AAnim
                If .NumFrames = 1 Then
                    MsgBox "Can't intrpolate animations with a single frame", vbOKOnly, "Interpolation error"
                    Exit Sub
                End If
                
                'Create new frames
                .NumFrames = .NumFrames * (num_interpolated_frames + 1) - frame_offset
                ReDim Preserve .Frames(.NumFrames - 1)
                'Move the original frames into their new positions
                For fi = .NumFrames - (1 + num_interpolated_frames - frame_offset) To 0 Step -next_elem_diff
                    .Frames(fi) = .Frames(fi / (num_interpolated_frames + 1))
                Next fi
                
                'Interpolate the new frames
                For fi = 0 To .NumFrames - (1 + next_elem_diff + num_interpolated_frames - frame_offset) Step next_elem_diff
                    For ifi = 1 To num_interpolated_frames
                        alpha = CSng(ifi) / CSng(num_interpolated_frames + 1)
                        GetTwoAFramesInterpolation hrc_sk, .Frames(fi), .Frames(fi + num_interpolated_frames + 1), alpha, .Frames(fi + ifi)
                    Next ifi
                Next fi
                
                If is_loop = vbYes Then
                    base_final_frame = .NumFrames - num_interpolated_frames - 1
                    For ifi = 1 To num_interpolated_frames
                        alpha = CSng(ifi) / CSng(num_interpolated_frames + 1)
                        GetTwoAFramesInterpolation hrc_sk, .Frames(base_final_frame), .Frames(0), alpha, .Frames(base_final_frame + ifi)
                    Next ifi
                End If
            End With
        Case K_AA_SKELETON:
            anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
            With DAAnims.BodyAnimations(anim_index)
                'Numframes1 and NumFrames2 are usually different. Don't know if this is relevant at all, but keep the balance between them just in case
                primary_secondary_counters_coef = .NumFrames1 / .NumFrames2

                If .NumFrames2 = 1 Then
                    MsgBox "Can't intrpolate animations with a single frame", vbOKOnly, "Interpolation error"
                    Exit Sub
                End If
                
                'Create new frames
                .NumFrames2 = .NumFrames2 * (num_interpolated_frames + 1) - frame_offset
                .NumFrames1 = .NumFrames2 * primary_secondary_counters_coef
                
                ReDim Preserve .Frames(.NumFrames2 - 1)
                'Move the original frames into their new positions
                For fi = .NumFrames2 - (1 + num_interpolated_frames - frame_offset) To 0 Step -next_elem_diff
                    .Frames(fi) = .Frames(fi / (num_interpolated_frames + 1))
                Next fi
                
                'Interpolate the new frames
                For fi = 0 To .NumFrames2 - (1 + next_elem_diff + num_interpolated_frames - frame_offset) Step next_elem_diff
                    For ifi = 1 To num_interpolated_frames
                        alpha = CSng(ifi) / CSng(num_interpolated_frames + 1)
                        If aa_sk.NumBones > 1 Then
                            GetTwoDAFramesInterpolation aa_sk, .Frames(fi), .Frames(fi + num_interpolated_frames + 1), alpha, .Frames(fi + ifi)
                        Else
                            GetTwoDAFramesWeaponInterpolation aa_sk, .Frames(fi), .Frames(fi + num_interpolated_frames + 1), alpha, .Frames(fi + ifi)
                        End If
                    Next ifi
                    'Interpolate the first frame too
                    'GetTwoDAFramesInterpolation AA_sk, .Frames(fi), .Frames(fi + num_interpolated_frames + 1), 0, aux_daframe
                    'CopyDAFrame aux_daframe, .Frames(fi)
                Next fi
                
                base_final_frame = .NumFrames2 - num_interpolated_frames - 1
                If is_loop = vbYes Then
                    For ifi = 1 To num_interpolated_frames
                        alpha = CSng(ifi) / CSng(num_interpolated_frames + 1)
                        If aa_sk.NumBones > 1 Then
                            GetTwoDAFramesInterpolation aa_sk, .Frames(base_final_frame), .Frames(0), alpha, .Frames(base_final_frame + ifi)
                        Else
                            GetTwoDAFramesWeaponInterpolation aa_sk, .Frames(base_final_frame), .Frames(0), alpha, .Frames(base_final_frame + ifi)
                        End If
                    Next ifi
                End If
                'GetTwoDAFramesInterpolation AA_sk, .Frames(base_final_frame), .Frames(.NumFrames2 - 1), 1#, aux_daframe
                'CopyDAFrame aux_daframe, .Frames(.NumFrames2 - 1)
                
                'NormalizeDAAnimationsPackAnimation DAAnims.BodyAnimations(anim_index)
            End With
                
            'Also don't forget about the weapon frames (where available)
            If anim_index < DAAnims.NumWeaponAnimations And aa_sk.NumWeapons > 0 Then
                With DAAnims.WeaponAnimations(anim_index)
                    .NumFrames2 = DAAnims.BodyAnimations(anim_index).NumFrames2
                    .NumFrames1 = DAAnims.BodyAnimations(anim_index).NumFrames1
                    
                    ReDim Preserve .Frames(.NumFrames2 - 1)
                    'Move the original frames into their new positions
                    For fi = .NumFrames2 - (1 + num_interpolated_frames - frame_offset) To 0 Step -next_elem_diff
                        .Frames(fi) = .Frames(fi / (num_interpolated_frames + 1))
                    Next fi
                    
                    'Interpolate the new frames
                    For fi = 0 To .NumFrames2 - (1 + num_interpolated_frames + num_interpolated_frames - frame_offset) Step next_elem_diff
                        For ifi = 1 To num_interpolated_frames
                            alpha = CSng(ifi) / CSng(num_interpolated_frames + 1)
                            GetTwoDAFramesWeaponInterpolation aa_sk, .Frames(fi), .Frames(fi + next_elem_diff), alpha, .Frames(fi + ifi)
                        Next ifi
                        'GetTwoDAFramesWeaponInterpolation AA_sk, .Frames(fi), .Frames(fi + num_interpolated_frames + 1), 0, aux_daframe
                        'CopyDAFrame aux_daframe, .Frames(fi)
                    Next fi
                    
                    base_final_frame = .NumFrames2 - num_interpolated_frames - 1
                    If is_loop = vbYes Then
                        For ifi = 1 To num_interpolated_frames
                            alpha = CSng(ifi) / CSng(num_interpolated_frames + 1)
                            GetTwoDAFramesWeaponInterpolation aa_sk, .Frames(base_final_frame), .Frames(0), alpha, .Frames(base_final_frame + ifi)
                        Next ifi
                    End If
                    'GetTwoDAFramesWeaponInterpolation AA_sk, .Frames(base_final_frame), .Frames(.NumFrames2 - 1), 1#, aux_daframe
                    'CopyDAFrame aux_daframe, .Frames(.NumFrames2 - 1)
                End With
                
                'NormalizeDAAnimationsPackAnimation DAAnims.WeaponAnimations(anim_index)
            End If
    End Select
    
    CurrentFrameScroll.max = CurrentFrameScroll.max * (num_interpolated_frames + 1) - frame_offset
    CurrentFrameScroll.value = CurrentFrameScroll.value * (num_interpolated_frames + 1)
    Call SetFrameEditorFields

    Picture1_Paint
End Sub

Private Sub InterpolateFrameButton_Click()
    Dim anim_index As Integer
    Dim fi As Integer
    Dim current_frame As Integer
    Dim next_frame As Integer
    Dim num_frames As Integer
    Dim num_interpolated_frames_str As String
    Dim num_interpolated_frames As Integer
    Dim alpha As Single
    Dim primary_secondary_counter_coef As Single
    
    num_interpolated_frames_str = InputBox("Number of frames to interpolate between each frame", "Animation interpolation", IIf(ModelType = K_HRC_SKELETON, "1", "2"))
    
    If num_interpolated_frames_str = "" Or Not IsNumeric(num_interpolated_frames_str) Then
        Exit Sub
    End If
    
    num_interpolated_frames = CInt(num_interpolated_frames_str)
    
    current_frame = CurrentFrameScroll.value
    next_frame = current_frame + num_interpolated_frames + 1
    
    If CurrentFrameScroll.value = CurrentFrameScroll.max - 1 Then
        next_frame = 0
    End If
    
    AddStateToBuffer
    
    Select Case ModelType
        Case K_HRC_SKELETON:
            With AAnim
                'Create new frames
                .NumFrames = .NumFrames + num_interpolated_frames
                ReDim Preserve .Frames(.NumFrames - 1)
                'Move the original frames into their new positions
                For fi = .NumFrames - 1 To current_frame + num_interpolated_frames Step -1
                    .Frames(fi) = .Frames(fi - num_interpolated_frames)
                Next fi
                
                'Interpolate the new frames
                For fi = 1 To num_interpolated_frames
                    alpha = CSng(fi) / CSng(num_interpolated_frames + 1)
                    GetTwoAFramesInterpolation hrc_sk, .Frames(current_frame), .Frames(next_frame), alpha, .Frames(current_frame + fi)
                Next fi
            End With
        Case K_AA_SKELETON:
            anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
            With DAAnims.BodyAnimations(anim_index)
                primary_secondary_counter_coef = .NumFrames1 / .NumFrames2
                'Create new frames
                .NumFrames2 = .NumFrames2 + num_interpolated_frames
                .NumFrames1 = .NumFrames2 * primary_secondary_counter_coef
                ReDim Preserve .Frames(.NumFrames2 - 1)
                'Move the original frames into their new positions
                For fi = .NumFrames2 - 1 To current_frame + num_interpolated_frames + 1 Step -1
                    .Frames(fi) = .Frames(fi - num_interpolated_frames)
                Next fi
                
                'Interpolate the new frames
                For fi = 1 To num_interpolated_frames
                    alpha = CSng(fi) / CSng(num_interpolated_frames + 1)
                    GetTwoDAFramesInterpolation aa_sk, .Frames(current_frame), .Frames(next_frame), alpha, .Frames(current_frame + fi)
                Next fi
            End With
                
            'Also don't forget about the weapon frames (where available)
            If anim_index < DAAnims.NumWeaponAnimations And aa_sk.NumWeapons > 0 Then
                num_frames = DAAnims.BodyAnimations(anim_index).NumFrames2
                With DAAnims.WeaponAnimations(anim_index)
                    'Create new frames
                    .NumFrames2 = .NumFrames2 + num_interpolated_frames
                    .NumFrames1 = .NumFrames2 * primary_secondary_counter_coef
                    ReDim Preserve .Frames(.NumFrames2 - 1)
                    'Move the original frames into their new positions
                    For fi = .NumFrames2 - 1 To current_frame + num_interpolated_frames + 1 Step -1
                        .Frames(fi) = .Frames(fi - num_interpolated_frames)
                    Next fi
                    
                    'Interpolate the new frames
                    For fi = 1 To num_interpolated_frames
                        alpha = CSng(fi) / CSng(num_interpolated_frames + 1)
                        GetTwoDAFramesWeaponInterpolation aa_sk, .Frames(current_frame), .Frames(next_frame), alpha, .Frames(current_frame + fi)
                    Next fi
                End With
            End If
    End Select

    CurrentFrameScroll.max = CurrentFrameScroll.max + num_interpolated_frames
    Call SetFrameEditorFields
    Picture1_Paint
    'CurrentFrameScroll.max = CurrentFrameScroll.max + num_interpolated_frames
    'CurrentFrameScroll.value = CurrentFrameScroll.value + 1
    'Call SetFrameEditorFields
    'Picture1_Paint
End Sub

Private Sub LightPosXScroll_Change()
    Call Picture1_Paint
End Sub

Private Sub LightPosYScroll_Change()
    Call Picture1_Paint
End Sub

Private Sub LightPosZScroll_Change()
    Call Picture1_Paint
End Sub

Private Sub MoveTexureUpDown_DownClick()
    Dim tex_index As Integer
    Dim tex_id_temp As Long
    Dim temp_tex As TEXTexture
    Select Case ModelType
        Case K_HRC_SKELETON:
            If SelectedBone > -1 Then
                If hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).NumTextures > 0 Then
                    tex_index = TextureSelectCombo.ListIndex
                    With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece)
                        If (tex_index < .NumTextures - 1) Then
                            temp_tex = .textures(tex_index)
                            .textures(tex_index) = .textures(tex_index + 1)
                            .textures(tex_index + 1) = temp_tex
                            SetTextureEditorFields
                            TextureSelectCombo.ListIndex = tex_index + 1
                        Else
                            Beep
                        End If
                    End With
                End If
            End If
        Case K_AA_SKELETON:
            tex_index = TextureSelectCombo.ListIndex
            With aa_sk
                If (tex_index < .NumTextures - 1) Then
                    temp_tex = .textures(tex_index)
                    .textures(tex_index) = .textures(tex_index + 1)
                    .textures(tex_index + 1) = temp_tex
                    tex_id_temp = .TexIDS(tex_index)
                    .TexIDS(tex_index) = .TexIDS(tex_index + 1)
                    .TexIDS(tex_index + 1) = tex_id_temp
                    SetTextureEditorFields
                    TextureSelectCombo.ListIndex = tex_index + 1
                Else
                    Beep
                End If
            End With
    End Select
    Picture1_Paint
End Sub

Private Sub MoveTexureUpDown_UpClick()
    Dim tex_index As Integer
    Dim tex_id_temp As Long
    Dim temp_tex As TEXTexture
    Select Case ModelType
        Case K_HRC_SKELETON:
            If SelectedBone > -1 Then
                If hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).NumTextures > 0 Then
                    tex_index = TextureSelectCombo.ListIndex
                    If (tex_index > 0) Then
                        With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece)
                            temp_tex = .textures(tex_index)
                            .textures(tex_index) = .textures(tex_index - 1)
                            .textures(tex_index - 1) = temp_tex
                            SetTextureEditorFields
                            TextureSelectCombo.ListIndex = tex_index - 1
                        End With
                    Else
                        Beep
                    End If
                End If
            End If
        Case K_AA_SKELETON:
            tex_index = TextureSelectCombo.ListIndex
            If (tex_index > 0) Then
                With aa_sk
                    temp_tex = .textures(tex_index)
                    .textures(tex_index) = .textures(tex_index - 1)
                    .textures(tex_index - 1) = temp_tex
                    tex_id_temp = .TexIDS(tex_index)
                    .TexIDS(tex_index) = .TexIDS(tex_index - 1)
                    .TexIDS(tex_index - 1) = tex_id_temp
                    SetTextureEditorFields
                    TextureSelectCombo.ListIndex = tex_index - 1
                End With
            Else
                Beep
            End If
    End Select
    Picture1_Paint
End Sub


Private Sub RemoveFrameButton_Click()
    Dim anim_index As Integer
    Dim fi As Integer
    Dim primary_secondary_counter_coef As Single
    
    AddStateToBuffer
    Select Case ModelType
        Case K_HRC_SKELETON:
            If RemoveFrameAAnimation(AAnim, CurrentFrameScroll.value) Then
                If (CurrentFrameScroll.value = CurrentFrameScroll.max - 1) Then _
                    CurrentFrameScroll.value = CurrentFrameScroll.value - 1
                CurrentFrameScroll.max = CurrentFrameScroll.max - 1
            Else
                Beep
            End If
        Case K_AA_SKELETON:
            anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
            With DAAnims.BodyAnimations(anim_index)
                If .NumFrames2 > 1 Then
                    For fi = CurrentFrameScroll.value To .NumFrames2 - 2
                        .Frames(fi) = .Frames(fi + 1)
                    Next fi
                    primary_secondary_counter_coef = .NumFrames1 / .NumFrames2
                    .NumFrames2 = .NumFrames2 - 1
                    .NumFrames1 = .NumFrames2 * primary_secondary_counter_coef
                    ReDim Preserve .Frames(.NumFrames2 - 1)
                    
                    'Also don't forget the weapon frames if available
                    If anim_index < DAAnims.NumWeaponAnimations And aa_sk.NumWeapons > 0 Then
                        For fi = CurrentFrameScroll.value To .NumFrames2 - 2
                            DAAnims.WeaponAnimations(anim_index).Frames(fi) = _
                                DAAnims.WeaponAnimations(anim_index).Frames(fi + 1)
                        Next fi
                        ReDim Preserve DAAnims.WeaponAnimations(anim_index).Frames(.NumFrames2 - 1)
                        DAAnims.WeaponAnimations(anim_index).NumFrames2 = .NumFrames2
                        DAAnims.WeaponAnimations(anim_index).NumFrames1 = .NumFrames1
                    End If
                    
                    If (CurrentFrameScroll.value = CurrentFrameScroll.max - 1) Then _
                        CurrentFrameScroll.value = CurrentFrameScroll.value - 1
                    CurrentFrameScroll.max = CurrentFrameScroll.max - 1
                Else
                    Beep
                End If
            End With
    End Select

    Picture1_Paint
End Sub

Private Sub RemoveTextureButton_Click()
    Dim ti As Integer
    Dim tex_index As Integer
    Select Case ModelType
        Case K_HRC_SKELETON:
            If SelectedBone > -1 Then
                AddStateToBuffer
                
                If hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).NumTextures > 0 Then
                    tex_index = TextureSelectCombo.ListIndex
                    With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece)
                        'This is dirty, but will prevent problems with the undo/redo
                        'UnloadTexture .textures(tex_index)
                        For ti = tex_index To .NumTextures - 2
                            .textures(ti) = .textures(ti + 1)
                        Next ti
                        ReDim Preserve .textures(IIf(.NumTextures = 1, 0, .NumTextures - 2))
                        .NumTextures = .NumTextures - 1
                        SetTextureEditorFields
                    End With
                Else
                    Beep
                End If
            End If
        Case K_AA_SKELETON:
            AddStateToBuffer
            
            tex_index = TextureSelectCombo.ListIndex
            With aa_sk
                If .NumTextures > 0 Then
                    'This is dirty, but will prevent problems with the undo/redo
                    'UnloadTexture .textures(tex_index)
                    For ti = tex_index To .NumTextures - 2
                        .textures(ti) = .textures(ti + 1)
                        .TexIDS(ti) = .TexIDS(ti + 1)
                    Next ti
                    ReDim Preserve .textures(IIf(.NumTextures = 1, 0, .NumTextures - 2))
                    ReDim Preserve .TexIDS(IIf(.NumTextures = 1, 0, .NumTextures - 2))
                    .NumTextures = .NumTextures - 1
                    If .NumTextures = 0 Then .TexIDS(0) = 0
                    SetTextureEditorFields
                Else
                    Beep
                End If
            End With
    End Select
    Picture1_Paint
End Sub

Private Sub SaveFF7AnimationButton_Click()
    Dim pattern As String
    Dim error_message As String
           
    On Error GoTo hand
    
    Picture1.Enabled = False
        
    If loaded Then
        Select Case (ModelType)
            Case K_HRC_SKELETON:
                pattern = "FF7 field animation (*.A)|*.a"
                CommonDialog1.Filter = pattern
                CommonDialog1.CancelError = True
                CommonDialog1.ShowSave 'Display the Open File Common Dialog
                
                If (CommonDialog1.fileName <> "") Then
                    If FileExist(CommonDialog1.fileName) Then _
                        If MsgBox("File already exists, overwrite it?", vbYesNo, "Confirmation") = vbNo Then _
                            GoTo cleanup
                            
                    WriteAAnimation AAnim, AAnim.AFile
                End If
            Case K_AA_SKELETON:
                If aa_sk.IsLimitBreak Then
                    pattern = "FF7 limit break animations pack (*.a00)|*.a00"
                Else
                    pattern = "FF7 battle animations pack (*DA)|*da"
                End If
                CommonDialog1.Filter = pattern
                CommonDialog1.CancelError = True
                CommonDialog1.ShowSave 'Display the Open File Common Dialog
                
                If (CommonDialog1.fileName <> "") Then
                    If FileExist(CommonDialog1.fileName) Then _
                        If MsgBox("File already exists, overwrite it?", vbYesNo, "Confirmation") = vbNo Then _
                            GoTo cleanup
                            
                    WriteDAAnimationsPack CommonDialog1.fileName, DAAnims
                End If
        End Select
    End If
    GoTo cleanup
hand:
    If Err <> 32755 Then
        error_message = "Error" + Str(Err)
        MsgBox error_message, vbOKOnly, "Unknow error Saving"
    End If
cleanup:
    Timer1.Enabled = True
End Sub

Private Sub ShowCharModelDBButton_Click()
    FF7FieldDBForm.FillControls
    FF7FieldDBForm.Show
End Sub

Private Sub ShowGroundCheck_Click()
    Picture1_Paint
End Sub

Private Sub ShowLastFrameGhostCheck_Click()
    Picture1_Paint
End Sub

Private Sub ShowTextureOptionsButton_Click()
    If ShowTextureOptionsButton.Caption = "Show Texture options" Then
        ShowTextureOptionsButton.Caption = "Show bone options"
        SelectedBoneFrame.Left = Me.ScaleWidth + 20
        TexturesFrame.Left = Me.ScaleWidth - 115
    Else
        ShowTextureOptionsButton.Caption = "Show Texture options"
        TexturesFrame.Left = Me.ScaleWidth + 20
        SelectedBoneFrame.Left = Me.ScaleWidth - 115
    End If
End Sub

Private Sub TextureSelectCombo_Change()
    Call TextureViewer_Paint
End Sub

Private Sub TextureSelectCombo_Click()
    Call TextureViewer_Paint
End Sub

Private Sub TextureViewer_Paint()
    If TextureSelectCombo.ListIndex > -1 Then
        Dim ti As Integer
        Dim aux As POINTAPI
        SetStretchBltMode TextureViewer.hdc, HALFTONE
        SetBrushOrgEx TextureViewer.hdc, 0, 0, aux
        TextureViewer.Cls
        
        Select Case ModelType
            Case K_HRC_SKELETON:
                If SelectedBone > -1 And SelectedBonePiece > -1 Then
                    If hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).textures(TextureSelectCombo.ListIndex).tex_id <> -1 Then
                        With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).textures(TextureSelectCombo.ListIndex)
                            ZeroAsTransparent.Enabled = True
                            ZeroAsTransparent.value = IIf(.ColorKeyFlag = 1, vbChecked, vbUnchecked)
                            StretchBlt TextureViewer.hdc, 0, 0, TextureViewer.ScaleWidth, TextureViewer.ScaleHeight, _
                                .hdc, 0, 0, .width, .height, vbSrcCopy
                        End With
                    Else
                        TextureViewer.Print "Texture not loaded"
                        ZeroAsTransparent.Enabled = False
                    End If
                End If
            Case K_AA_SKELETON:
                If aa_sk.textures(TextureSelectCombo.ListIndex).tex_id <> -1 Then
                    With aa_sk.textures(TextureSelectCombo.ListIndex)
                        ZeroAsTransparent.Enabled = True
                        ZeroAsTransparent.value = IIf(.ColorKeyFlag = 1, vbChecked, vbUnchecked)
                        StretchBlt TextureViewer.hdc, 0, 0, TextureViewer.ScaleWidth, TextureViewer.ScaleHeight, _
                            .hdc, 0, 0, .width, .height, vbSrcCopy
                    End With
                Else
                    TextureViewer.Print "Texture not loaded"
                    ZeroAsTransparent.Enabled = False
                End If
            Case Else
                With TextureViewer
                    TextureViewer.Print "Texture not loaded"
                    ZeroAsTransparent.Enabled = False
                    BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, vbWhiteness
                End With
        End Select
    Else
        With TextureViewer
            ZeroAsTransparent.Enabled = False
            BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, Whiteness
        End With
    End If
End Sub


Private Sub Timer1_Timer()
    Sleep 100
    Timer1.Enabled = False
    Picture1.Enabled = True
End Sub

Private Sub XAnimationFramePartText_Change()
    If LoadingAnimationQ Then _
        Exit Sub
        
    If IsNumeric(XAnimationFramePartText.Text) Then
        If XAnimationFramePartText.Text * 10000 <= XAnimationFramePartUpDown.max _
        And XAnimationFramePartText.Text * 10000 >= XAnimationFramePartUpDown.Min Then
            XAnimationFramePartUpDown.value = XAnimationFramePartText.Text * 10000
            'Call XAnimationFramePartUpDown_Change
        End If
    Else
        Beep
    End If
End Sub

Private Sub YAnimationFramePartText_Change()
    If LoadingAnimationQ Then _
        Exit Sub
        
    If IsNumeric(YAnimationFramePartText.Text) Then
        If YAnimationFramePartText.Text * 10000 <= YAnimationFramePartUpDown.max _
        And YAnimationFramePartText.Text * 10000 >= YAnimationFramePartUpDown.Min Then
            YAnimationFramePartUpDown.value = YAnimationFramePartText.Text * 10000
            'Call YAnimationFramePartUpDown_Change
        End If
    Else
        Beep
    End If
End Sub

Private Sub ZAnimationFramePartText_Change()
    If LoadingAnimationQ Then _
        Exit Sub
        
    If IsNumeric(ZAnimationFramePartText.Text) Then
        If ZAnimationFramePartText.Text * 10000 <= ZAnimationFramePartUpDown.max _
        And ZAnimationFramePartText.Text * 10000 >= ZAnimationFramePartUpDown.Min Then
            ZAnimationFramePartUpDown.value = ZAnimationFramePartText.Text * 10000
            'Call ZAnimationFramePartUpDown_Change
        End If
    Else
        Beep
    End If
End Sub

Private Sub FrontLight_Click()
    Picture1_Paint
End Sub

Private Sub RearLight_Click()
    Picture1_Paint
End Sub

Private Sub LeftLight_Click()
    Picture1_Paint
End Sub

Private Sub RemovePieceButton_Click()
    On Error GoTo hand
    
    If ModelType = K_HRC_SKELETON Then
        If hrc_sk.Bones(SelectedBone).NumResources > 0 Then
            AddStateToBuffer
            RemoveHRCBonePiece hrc_sk.Bones(SelectedBone), SelectedBonePiece
        End If
    Else
        If aa_sk.Bones(SelectedBone).NumModels > 0 Then
            AddStateToBuffer
            RemoveAABoneModel aa_sk.Bones(SelectedBone), SelectedBonePiece
        End If
    End If
    SelectedBonePiece = -1
    SelectedPieceFrame.Enabled = False
    Picture1_Paint
    Exit Sub
hand:
    If Err <> 32755 Then
        MsgBox "Error" + Str$(Err), vbOKOnly, "Unknow error Removing"
    End If
End Sub


Private Sub ResizeBoneXUpDown_Change()
    If LoadingBoneModifiersQ Then _
        Exit Sub
        
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    Select Case ModelType
        Case K_HRC_SKELETON:
            hrc_sk.Bones(SelectedBone).ResizeX = ResizeBoneXUpDown.value / 100
        Case K_AA_SKELETON:
            If (SelectedBone = aa_sk.NumBones) Then
                aa_sk.WeaponModels(WeaponCombo.ListIndex).ResizeX = ResizeBoneXUpDown.value / 100
            Else
                aa_sk.Bones(SelectedBone).ResizeX = ResizeBoneXUpDown.value / 100
            End If
    End Select
    Picture1_Paint
    DoNotAddStateQ = False
End Sub



Private Sub RightLight_Click()
Picture1_Paint
End Sub

Private Sub BoneSelector_Click()
    If loaded Then
        SelectedBone = BoneSelector.ListIndex
        SelectedBonePiece = -1
        If SelectedBone > -1 Then
            SetBoneModifiers
            SelectedBoneFrame.Enabled = True
            If (ModelType = K_AA_SKELETON) Then _
                SetTextureEditorFields
        Else
            SelectedBoneFrame.Enabled = False
        End If
        Picture1_Paint
    End If
End Sub
Private Sub OpenFF7ModelButton_Click()
    On Error GoTo hand
    Dim pattern As String
    
    Picture1.Enabled = False
    pattern = "Any FF7 3D Model|*.*|FF7 Field Model file|*.p|FF7 Battle Model file|*.*|FF7 Field Skeleton file|*.hrc|FF7 Battle Skeleton file|*aa|FF7 Magic Skeleton file|*.d|3D Studio model|*.3ds"
    CommonDialog1.Filter = pattern
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen 'Display the Open File Common Dialog
   
    If (CommonDialog1.fileName <> "") Then
        If loaded Then DestroyCurrentModel
        OpenFF7File CommonDialog1.fileName
    End If
    
    Editor.Hide
    Timer1.Enabled = True
    Exit Sub
hand:
    If Err <> 32755 Then
        MsgBox "Error" + Str$(Err), vbOKOnly, "Unknow error Loading"
    End If
    Timer1.Enabled = True
End Sub

Private Sub ChangeAnimationButton_Click()
    On Error GoTo hand
    
    Picture1.Enabled = False
    If loaded Then
        Dim pattern As String
        If ModelType = K_AA_SKELETON Then
            pattern = GetModelAnimationPacksFilter(aa_sk.fileName)
        Else
            pattern = "FF7 Field-model animation|*.a"
        End If
        CommonDialog1.Filter = pattern
        CommonDialog1.CancelError = True
        CommonDialog1.ShowOpen 'Display the Open File Common Dialog
        
        If ModelType = K_AA_SKELETON Then
            SetBattleModelAnimationsPack CommonDialog1.fileName
        Else
            SetFieldModelAnimation CommonDialog1.fileName
        End If
    End If
    Timer1.Enabled = True
Exit Sub
hand:
    If Err <> 32755 Then
        MsgBox "Error" + Str(Err), vbOKOnly, "Unknow error loading animation"
    End If
    Timer1.Enabled = True
End Sub
Private Sub SaveFF7ModelButton_Click()
    Dim p_min As Point3D
    Dim p_max As Point3D
    Dim pattern As String
    Dim bone_i As Integer
    Dim part_i As Integer
    Dim vert_i As Integer
    Dim group_i As Integer
    Dim fin As Integer
    Dim mes As String
    Dim write_allowed As Boolean
    Dim C As color
    Dim jsp As Integer
    Dim anims_pack_filename As String
    Dim aux_weapon_anim As DAFrame
    Dim anim_index As Integer
           
    On Error GoTo hand
    
    Picture1.Enabled = False
        
    If loaded Then
        Select Case (ModelType)
            Case K_HRC_SKELETON:
                pattern = "FF7 HRC file (field skeleton)|*.hrc"
                CommonDialog1.Filter = pattern
                CommonDialog1.CancelError = True
                CommonDialog1.ShowSave 'Display the Open File Common Dialog
                
                If (CommonDialog1.fileName <> "") Then
                    If FileExist(CommonDialog1.fileName) Then _
                        If MsgBox("File already exists, overwrite (including asociated RSB & P files)?", vbYesNo, "Confirmation") = vbNo Then _
                            GoTo cleanup
                        
                    AddStateToBuffer
                    
                    ComputeHRCBoundingBox hrc_sk, AAnim.Frames(CurrentFrameScroll.value), p_min, p_max
                    SetCameraAroundModel p_min, p_max, 0, 0, -2 * ComputeSceneRadius(p_min, p_max), _
                        0, 0, 0, 1, 1, 1
                    
                    SetLights
                    
                    ApplyHRCChanges hrc_sk, AAnim.Frames(CurrentFrameScroll.value), _
                                    (MsgBox("Compile multi-P bones in a single file?", vbYesNo, "Confirmation") = vbYes)
                    
                    WriteHRCSkeleton hrc_sk, CommonDialog1.fileName
                    'WriteAAnimation AAnim, AAnim.AFile
                    CreateDListsFromHRCSkeleton hrc_sk
                End If
            Case K_AA_SKELETON:
                pattern = "FF7 AA file (battle skeleton)|*aa"
                CommonDialog1.Filter = pattern
                CommonDialog1.CancelError = True
                CommonDialog1.ShowSave 'Display the Open File Common Dialog
                
                If (CommonDialog1.fileName <> "") Then
                    If FileExist(CommonDialog1.fileName) Then _
                        If MsgBox("File already exists, overwrite (including asociated files)?", vbYesNo, "Confirmation") = vbNo Then _
                            GoTo cleanup
                            
                    AddStateToBuffer
                    
                    If Not aa_sk.IsBattleLocation Then _
                            anim_index = val(BattleAnimationCombo.List(BattleAnimationCombo.ListIndex))
                    ComputeAABoundingBox aa_sk, _
                        DAAnims.BodyAnimations(anim_index).Frames(CurrentFrameScroll.value), _
                        p_min, p_max
                    SetCameraAroundModel p_min, p_max, 0, 0, -2 * ComputeSceneRadius(p_min, p_max), _
                        0, 0, 0, 1, 1, 1
                        
                    SetLights
                
                    If (aa_sk.NumWeaponAnims > 0) Then _
                        aux_weapon_anim = DAAnims.WeaponAnimations(0).Frames(0)
                    
                    ApplyAAChanges aa_sk, DAAnims.BodyAnimations(0).Frames(0), aux_weapon_anim 'DAAnims.WeaponAnimations(0).Frames(0)
                    
                    If LCase(Right$(CommonDialog1.fileName, 2)) = ".d" Then
                        'Magic model
                        'If aa_sk.IsLimitBreak Then
                        '    If MsgBox("The skeleton was loaded as a Limit Break. Overwrite it? Choose 'No' to write just animations.", vbYesNo, "Confirmation") = vbYes Then _
                        '        WriteAASkeleton BATTLE_LGP_PATH + "\" + GetLimitCharacterFileName(CommonDialog1.filename), aa_sk
                        'Else
                            WriteMagicSkeleton CommonDialog1.fileName, aa_sk
                        'End If
                        'anims_pack_filename = Left$(CommonDialog1.filename, _
                        '    Len(CommonDialog1.filename) - 2) + ".a00"
                    Else
                        'Standard battle model
                        WriteAASkeleton CommonDialog1.fileName, aa_sk
                        'anims_pack_filename = Left$(CommonDialog1.filename, _
                        '    Len(CommonDialog1.filename) - 2) + "da"
                    End If
                    'WriteDAAnimationsPack anims_pack_filename, DAAnims
                    'CheckWriteDAAnimationsPack anims_pack_filename, DAAnims
                    CreateDListsFromAASkeleton aa_sk
                End If
            Case Else:
                pattern = "FF7 Field Model file|*.p|FF7 Battle Model file|*.*"
                CommonDialog1.Filter = pattern
                CommonDialog1.CancelError = True
                CommonDialog1.ShowSave 'Display the Open File Common Dialog
                
                If (CommonDialog1.fileName <> "") Then
                    If FileExist(CommonDialog1.fileName) Then _
                        If MsgBox("File already exists, overwrite?", vbYesNo, "Confirmation") = vbNo Then _
                            GoTo cleanup
                    
                    AddStateToBuffer
                    
                    glMatrixMode GL_MODELVIEW
                    glPushMatrix
                    With P_Model
                        SetCameraModelViewQuat .RepositionX, .RepositionY, _
                                        .RepositionZ, _
                                        .RotationQuaternion, _
                                        .ResizeX, .ResizeY, .ResizeZ
                    End With
                    ApplyPChanges P_Model, (LCase(Right$(CommonDialog1.fileName, 2)) <> ".p")
                    
                    ComputePModelBoundingBox P_Model, p_min, p_max
                    SetCameraAroundModel p_min, p_max, _
                        0, 0, -2 * ComputeSceneRadius(p_min, p_max), _
                        0, 0, 0, 1, 1, 1
                    SetLights
                    
                    If glIsEnabled(GL_LIGHTING) = GL_TRUE Then _
                        ApplyCurrentVColors P_Model
                        
                    glPopMatrix
                    WritePModel P_Model, CommonDialog1.fileName
                    CreateDListsFromPModel P_Model
                End If
        End Select
        FrontLight.value = vbUnchecked
        RearLight.value = vbUnchecked
        RightLight.value = vbUnchecked
        LeftLight.value = vbUnchecked
    End If
    GoTo cleanup
Exit Sub
hand:
    If Err <> 32755 Then
        mes = "Error" + Str(Err)
        MsgBox mes, vbOKOnly, "Unknow error Saving"
    End If
cleanup:
    Timer1.Enabled = True
End Sub




Private Sub Form_Load()
    UNDO_BUFFER_CAPACITY = 30
    ReadCFGFile
    ReadCharFilterFile
    ReDim UnDoBuffer(UNDO_BUFFER_CAPACITY)
    ReDim ReDoBuffer(UNDO_BUFFER_CAPACITY)
    
    DoNotAddStateQ = False
    LightPosXScroll.max = LIGHT_STEPS
    LightPosXScroll.Min = -LIGHT_STEPS
    LightPosYScroll.max = LIGHT_STEPS
    LightPosYScroll.Min = -LIGHT_STEPS
    LightPosZScroll.max = LIGHT_STEPS
    LightPosZScroll.Min = -LIGHT_STEPS
    
    loaded = False
    BoneSelector.AddItem "None", 0
    UnDoCursor = 0
    ReDoCursor = 0
    
    ''Debug.Print "RES=", GetDeviceCaps(GetDC(0), HORZRES), GetDeviceCaps(GetDC(0), VERTRES)
    
    glEnable GL_DEPTH_TEST
    
    glClearColor 0.5, 0.5, 1, 0
    
    Set Editor = Forms.Add("PEditor")
    Dim comm As String
    
    comm = GetCommLine
    
    If Left$(comm, 1) = Chr$(34) Then comm = Right$(comm, Len(comm) - 1)
    
    If Len(comm) > 0 Then OpenFF7File comm
    
    MinFormWidth = Me.width
    MinFormHeight = Me.height
    
    DontRefreshPicture = False
End Sub

Private Sub Form_Resize()
    With Me
        'The window can't be resize while minimized
        If .WindowState = 1 Then _
            Exit Sub
        If .width < MinFormWidth Then _
            .width = MinFormWidth
        If .height < MinFormHeight Then _
            .height = MinFormHeight
        If .ScaleWidth > 0 Then
            Picture1.width = .ScaleWidth - 120 - SelectedPieceFrame.width - 10
            Picture1.height = .ScaleHeight - 25
            OpenFF7ModelButton.Left = .ScaleWidth - 115
            ChangeAnimationButton.Left = .ScaleWidth - 115
            ShowCharModelDBButton.Left = .ScaleWidth - 115
            SaveFF7ModelButton.Left = .ScaleWidth - 115
            SaveFF7AnimationButton.Left = .ScaleWidth - 115
            SelectedBoneFrame.Left = .ScaleWidth - 115
            GeneralLightingFrame.Left = .ScaleWidth - 115
            ShowBonesCheck.Top = .ScaleHeight - 20
            BoneSelector.Left = .ScaleWidth - 115
            BoneSelectorLabel.Left = .ScaleWidth - 115
            AnimationFrameLabel.Top = .ScaleHeight - 20
            AnimationFrameText.Top = .ScaleHeight - 20
            CurrentFrameScroll.Top = .ScaleHeight - 20
            If AnimationOptionsButton.Caption = "Show Frame options" Then
               AnimationOptionsFrame.Left = .ScaleWidth + 10
            Else
                AnimationOptionsFrame.Left = .ScaleWidth - 115
            End If
            If ShowTextureOptionsButton.Caption = "Show Texture options" Then
                TexturesFrame.Left = .ScaleWidth + 10
            Else
                TexturesFrame.Left = .ScaleWidth - 115
            End If
            ShowTextureOptionsButton.Left = .ScaleWidth - 115
            AnimationOptionsButton.Left = .ScaleWidth - 115
            InterpolateAnimationButton.Left = .ScaleWidth - 115
            InterpolateAllAnimsCommand.Left = .ScaleWidth - 115
            Picture1_Paint
        End If
    End With
End Sub

Private Sub Form_Terminate()
    DisableOpenGL OGLContext
    Unload Editor
    Unload FF7FieldDBForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Editor
    Unload FF7FieldDBForm
    Unload InterpolateAllAnimsForm
End Sub

Private Sub CurrentFrameScroll_Change()
    If CurrentFrameScroll.value = CurrentFrameScroll.max Then CurrentFrameScroll.value = 0
    If CurrentFrameScroll.value = -1 Then CurrentFrameScroll.value = CurrentFrameScroll.max - 1
    AnimationFrameText.Text = CurrentFrameScroll.value
    Call SetFrameEditorFields
    Picture1_Paint
End Sub

Private Sub ResizeBoneYUpDown_Change()
    If LoadingBoneModifiersQ Then _
        Exit Sub
    
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    Select Case ModelType
        Case K_HRC_SKELETON:
            hrc_sk.Bones(SelectedBone).ResizeY = ResizeBoneYUpDown.value / 100
        Case K_AA_SKELETON:
            If (SelectedBone = aa_sk.NumBones) Then
                aa_sk.WeaponModels(WeaponCombo.ListIndex).ResizeY = ResizeBoneYUpDown.value / 100
            Else
                aa_sk.Bones(SelectedBone).ResizeY = ResizeBoneYUpDown.value / 100
            End If
    End Select
    Picture1_Paint
    DoNotAddStateQ = False
End Sub

Private Sub ResizeBoneZUpDown_Change()
    If LoadingBoneModifiersQ Then _
        Exit Sub

    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    Select Case ModelType
        Case K_HRC_SKELETON:
            hrc_sk.Bones(SelectedBone).ResizeZ = ResizeBoneZUpDown.value / 100
        Case K_AA_SKELETON:
            If (SelectedBone = aa_sk.NumBones) Then
                aa_sk.WeaponModels(WeaponCombo.ListIndex).ResizeZ = ResizeBoneZUpDown.value / 100
            Else
                aa_sk.Bones(SelectedBone).ResizeZ = ResizeBoneZUpDown.value / 100
            End If
    End Select
    Picture1_Paint
    DoNotAddStateQ = False
End Sub
Private Sub Picture1_DblClick()
    If loaded Then
        Select Case (ModelType)
            Case K_HRC_SKELETON:
                If SelectedBonePiece > -1 Then
                    EditedPModel = hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).Model
                    Editor.OpenP ""
                    EditedBone = SelectedBone
                    EditedBonePiece = SelectedBonePiece
                    Editor.Show
                End If
            Case K_AA_SKELETON:
                If SelectedBonePiece > -1 Then
                    EditedPModel = aa_sk.Bones(SelectedBone).Models(SelectedBonePiece)
                    Editor.OpenP ""
                    EditedBone = SelectedBone
                    EditedBonePiece = SelectedBonePiece
                    Editor.Show
                ElseIf SelectedBone = aa_sk.NumBones Then
                    EditedPModel = aa_sk.WeaponModels(WeaponCombo.List(WeaponCombo.ListIndex))
                    Editor.OpenP ""
                    EditedBone = SelectedBone
                    EditedBonePiece = SelectedBonePiece
                    Editor.Show
                End If
            Case K_P_BATTLE_MODEL:
                EditedPModel = P_Model
                Editor.OpenP ""
                Editor.Show
            Case K_P_FIELD_MODEL:
                EditedPModel = P_Model
                Editor.OpenP ""
                Editor.Show
        End Select
    End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim p_min As Point3D
    Dim p_max As Point3D
    
    Dim BI As Integer
    Dim PI As Integer
    
    Dim anim_index As Integer
    Dim weapon_index As Integer
    Dim weapon_frame As DAFrame
    
    If loaded Then
        SetOGLSettings
        
        glClearColor 0.5, 0.5, 1, 0
        glViewport 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        
        Select Case (ModelType)
            Case K_HRC_SKELETON:
                ComputeHRCBoundingBox hrc_sk, AAnim.Frames(CurrentFrameScroll.value), p_min, p_max
                SetCameraAroundModel p_min, p_max, PanX, PanY, PanZ + DIST, _
                    alpha, Beta, Gamma, 1, 1, 1
                
                BI = GetClosestHRCBone(hrc_sk, AAnim.Frames(CurrentFrameScroll.value), x, y, DIST)
                SelectedBone = BI
                BoneSelector.ListIndex = BI
                If BI > -1 Then
                    PI = GetClosestHRCBonePiece(hrc_sk, AAnim.Frames(CurrentFrameScroll.value), BI, x, y, DIST)
                    
                    SelectedBonePiece = PI
                    If PI > -1 Then
                        SetBonePieceModifiers
                        SelectedPieceFrame.Enabled = True
                    Else
                        SelectedBoneFrame.Enabled = False
                    End If
                    SetBoneModifiers
                    SelectedBoneFrame.Enabled = True
                Else
                    SelectedBonePiece = -1
                    SelectedBoneFrame.Enabled = False
                    SelectedPieceFrame.Enabled = False
                End If
                SetTextureEditorFields
                
            Case K_AA_SKELETON:
                If Not aa_sk.IsBattleLocation Then _
                    anim_index = val(BattleAnimationCombo.List(BattleAnimationCombo.ListIndex))
                ComputeAABoundingBox aa_sk, _
                    DAAnims.BodyAnimations(anim_index).Frames(CurrentFrameScroll.value), _
                    p_min, p_max
                SetCameraAroundModel p_min, p_max, PanX, PanY, PanZ + DIST, _
                    alpha, Beta, Gamma, 1, 1, 1
                weapon_index = -1
                If anim_index < DAAnims.NumWeaponAnimations And aa_sk.NumWeapons > 0 Then
                    If WeaponCombo.ListIndex > -1 Then _
                        weapon_index = WeaponCombo.List(WeaponCombo.ListIndex)
                    weapon_frame = DAAnims.WeaponAnimations(anim_index).Frames(CurrentFrameScroll.value)
                End If
                
                BI = GetClosestAABone(aa_sk, DAAnims.BodyAnimations(anim_index).Frames(CurrentFrameScroll.value), weapon_frame, weapon_index, x, y, DIST)
                SelectedBone = BI
                If BI <= aa_sk.NumBones Then _
                    BoneSelector.ListIndex = BI
                If BI > -1 And BI < aa_sk.NumBones Then
                    If SelectBoneForWeaponAttachmentQ Then _
                        SetWeaponAnimationAttachedToBone (Button = vbRightButton)
                    PI = GetClosestAABoneModel(aa_sk, DAAnims.BodyAnimations(anim_index).Frames(CurrentFrameScroll.value), BI, x, y, DIST)
                    
                    SelectedBonePiece = PI
                    SetBoneModifiers
                    If PI > -1 Then
                        SetBonePieceModifiers
                        SelectedPieceFrame.Enabled = True
                    Else
                        SelectedPieceFrame.Enabled = False
                    End If
                    SelectedBoneFrame.Enabled = True
                ElseIf BI = aa_sk.NumBones Then
                    SetBoneModifiers
                    SelectedBonePiece = -2
                    SelectedBoneFrame.Enabled = True
                    SetBonePieceModifiers
                    SelectedPieceFrame.Enabled = True
                Else
                    SelectedBonePiece = -1
                    SelectedBoneFrame.Enabled = False
                    SelectedPieceFrame.Enabled = False
                End If
        End Select
        
        'AnimationOptionsFrame.Enabled = SelectedBoneFrame.Enabled
        
        Picture1_Paint
        x_last = x
        y_last = y
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim p_min As Point3D
    Dim p_max As Point3D
    
    Dim p_temp As Point3D
    Dim p_temp2 As Point3D
    Dim aux_alpha As Single
    Dim aux_y As Single
    Dim aux_dist As Single
    Dim anim_index As Integer
    Dim wasValidQ As Boolean
    
    If loaded And Button <> 0 Then
        If ShowGroundCheck.value = vbChecked Then
            SetCameraModelView PanX, PanY, PanZ + DIST, _
                alpha, Beta, Gamma, 1, 1, 1
            wasValidQ = Not IsCameraUnderGround
        Else
            wasValidQ = False
        End If
        
        Select Case (Button)
            Case vbLeftButton:
                Beta = (Beta + x - x_last) Mod 360
                aux_alpha = alpha
                alpha = (alpha + y - y_last) Mod 360
                SetCameraModelView PanX, PanY, PanZ + DIST, _
                    alpha, Beta, Gamma, 1, 1, 1
                If wasValidQ And IsCameraUnderGround Then _
                    alpha = aux_alpha
            Case vbRightButton:
                aux_dist = DIST
                DIST = DIST + (y - y_last) * diameter / 100
                
                SetCameraModelView PanX, PanY, PanZ + DIST, _
                    alpha, Beta, Gamma, 1, 1, 1
                If wasValidQ And IsCameraUnderGround Then _
                   DIST = aux_dist
            Case vbRightButton + vbLeftButton:
                Select Case (ModelType)
                    Case K_P_BATTLE_MODEL:
                        SetCameraPModel P_Model, 0, 0, DIST, 0, 0, 0, 1, 1, 1
                    Case K_P_FIELD_MODEL:
                        SetCameraPModel P_Model, 0, 0, DIST, 0, 0, 0, 1, 1, 1
                    Case K_HRC_SKELETON:
                        ComputeHRCBoundingBox hrc_sk, AAnim.Frames(CurrentFrameScroll.value), p_min, p_max
                        SetCameraAroundModel p_min, p_max, 0, 0, DIST, _
                            0, 0, 0, 1, 1, 1
                    Case K_AA_SKELETON:
                        If Not aa_sk.IsBattleLocation Then _
                            anim_index = val(BattleAnimationCombo.List(BattleAnimationCombo.ListIndex))
                        ComputeAABoundingBox aa_sk, _
                            DAAnims.BodyAnimations(anim_index).Frames(CurrentFrameScroll.value), _
                            p_min, p_max
                        SetCameraAroundModel p_min, p_max, 0, 0, DIST, _
                            0, 0, 0, 1, 1, 1
                End Select
                
                aux_y = PanY
                
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
                
                SetCameraModelView PanX, PanY, PanZ + DIST, _
                    alpha, Beta, Gamma, 1, 1, 1
                If wasValidQ And IsCameraUnderGround Then _
                    PanY = aux_y
        End Select
        
        x_last = x
        y_last = y

        Picture1_Paint
    End If
End Sub
Private Sub Picture1_Paint()
    If loaded And Not DontRefreshPicture Then
        'If SelectBoneForWeaponAttachmentQ Then _
        '    Picture1.Cls
        glViewport 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
        glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
        SetDefaultOGLRenderState
        DrawCurrentModel
        SwapBuffers Picture1.hdc
        If SelectBoneForWeaponAttachmentQ Then
            Picture1.CurrentX = 0
            Picture1.CurrentY = 0
            Picture1.Print "Please choose a bone to attach the weapon to"
        End If
    End If
End Sub

Private Sub ShowBonesCheck_Click()
    Picture1_Paint
End Sub

Private Sub ResizeBoneXText_Change()
    If LoadingBoneModifiersQ Then _
        Exit Sub
        
    If IsNumeric(ResizeBoneXText.Text) Then
        ResizeBoneXUpDown.value = ResizeBoneXText.Text
        'Call ResizeBoneXUpDown_Change
    Else
        Beep
    End If
End Sub

Private Sub ResizeBoneYText_Change()
    If LoadingBoneModifiersQ Then _
        Exit Sub
        
    If IsNumeric(ResizeBoneYText.Text) Then
        ResizeBoneYUpDown.value = ResizeBoneYText.Text
        'Call ResizeBoneYUpDown_Change
    Else
        Beep
    End If
End Sub

Private Sub ResizeBoneZText_Change()
    If LoadingBoneModifiersQ Then _
        Exit Sub
        
    If IsNumeric(ResizeBoneZText.Text) Then
        ResizeBoneZUpDown.value = ResizeBoneZText.Text
        'Call ResizeBoneZUpDown_Change
    Else
        Beep
    End If
End Sub

Private Sub BoneLengthUpDown_Change()
    If LoadingBoneModifiersQ Then _
        Exit Sub
        
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    Select Case ModelType
        Case K_HRC_SKELETON:
            hrc_sk.Bones(SelectedBone).length = BoneLengthUpDown.value / 10000
            BoneLengthText = BoneLengthUpDown.value / 10000
        Case K_AA_SKELETON:
            aa_sk.Bones(SelectedBone).length = BoneLengthUpDown.value / 10000
            BoneLengthText = BoneLengthUpDown.value / 10000
    End Select
    Picture1_Paint
    DoNotAddStateQ = False
End Sub
Sub DrawCurrentModel()
    Dim aux_anim As DAFrame
    Dim anim_index As Integer
    Dim weapon_index As Integer
    Dim rot_mat(16) As Double
    Dim p_min As Point3D
    Dim p_max As Point3D
    Dim diameter As Integer
    Dim turn_on_lightsQ As Boolean
    
    Select Case ModelType
        Case K_P_FIELD_MODEL:
            ComputePModelBoundingBox P_Model, p_min, p_max
            SetCameraAroundModel p_min, p_max, PanX, PanY, PanZ + DIST, _
                alpha, Beta, Gamma, 1, 1, 1
            
            If ShowGroundCheck.value = vbChecked Then
                glDisable GL_LIGHTING
                DrawGround
                DrawShadow p_min, p_max
            End If
            
            SetLights
            
            glMatrixMode GL_MODELVIEW
            glPushMatrix
            With P_Model
                glTranslatef .RepositionX, .RepositionY, .RepositionZ
                
                BuildRotationMatrixWithQuaternionsXYZ .RotateAlpha, .RotateBeta, _
                    .RotateGamma, rot_mat
                glMultMatrixd rot_mat(0)
                
                glScalef .ResizeX, .ResizeY, .ResizeZ
            End With
            
            DrawPModel P_Model, tex_ids, False
            glPopMatrix
        Case K_P_BATTLE_MODEL:
            ComputePModelBoundingBox P_Model, p_min, p_max
            SetCameraAroundModel p_min, p_max, PanX, PanY, PanZ + DIST, _
                alpha, Beta, Gamma, 1, 1, 1
            
            If ShowGroundCheck.value = vbChecked Then
                glDisable GL_LIGHTING
                DrawGround
                DrawShadow p_min, p_max
            End If
            
            SetLights
            
            glMatrixMode GL_MODELVIEW
            glPushMatrix
            With P_Model
                glTranslatef .RepositionX, .RepositionY, .RepositionZ
                
                BuildRotationMatrixWithQuaternionsXYZ .RotateAlpha, .RotateBeta, _
                    .RotateGamma, rot_mat
                glMultMatrixd rot_mat(0)
                
                glScalef .ResizeX, .ResizeY, .ResizeZ
            End With
            
            DrawPModel P_Model, tex_ids, False
            glPopMatrix
        Case K_HRC_SKELETON:
            ComputeHRCBoundingBox hrc_sk, AAnim.Frames(CurrentFrameScroll.value), p_min, p_max
            SetCameraAroundModel p_min, p_max, PanX, PanY, PanZ + DIST, _
                alpha, Beta, Gamma, 1, 1, 1
                
            If ShowGroundCheck.value = vbChecked Then
                glDisable GL_LIGHTING
                DrawGround
                DrawShadow p_min, p_max
            End If
            SetLights
            DrawHRCSkeleton hrc_sk, AAnim.Frames(CurrentFrameScroll.value), DListEnableCheck.value = vbChecked
            If ShowLastFrameGhostCheck.value = vbChecked Then
                glColorMask GL_TRUE, GL_TRUE, GL_FALSE, GL_TRUE
                If CurrentFrameScroll.value = 0 Then
                    DrawHRCSkeleton hrc_sk, AAnim.Frames(CurrentFrameScroll.max - 1), DListEnableCheck.value = vbChecked
                Else
                    DrawHRCSkeleton hrc_sk, AAnim.Frames(CurrentFrameScroll.value - 1), DListEnableCheck.value = vbChecked
                End If
                glColorMask GL_TRUE, GL_TRUE, GL_TRUE, GL_TRUE
            End If
            SelectHRCBoneAndPiece hrc_sk, AAnim.Frames(CurrentFrameScroll.value), SelectedBone, SelectedBonePiece
            
            If ShowBonesCheck.value = vbChecked Then
                glDisable GL_DEPTH_TEST
                glDisable GL_LIGHTING
                glColor3f 0, 1, 0
                DrawHRCSkeletonBones hrc_sk, AAnim.Frames(CurrentFrameScroll.value)
                glEnable GL_DEPTH_TEST
            End If
               
        Case K_AA_SKELETON:
            
            If Not aa_sk.IsBattleLocation Then _
                anim_index = val(BattleAnimationCombo.List(BattleAnimationCombo.ListIndex))
            ComputeAABoundingBox aa_sk, _
                DAAnims.BodyAnimations(anim_index).Frames(CurrentFrameScroll.value), _
                p_min, p_max
            SetCameraAroundModel p_min, p_max, PanX, PanY, PanZ + DIST, _
                alpha, Beta, Gamma, 1, 1, 1
                
            If ShowGroundCheck.value = vbChecked Then
                glDisable GL_LIGHTING
                DrawGround
                DrawShadow p_min, p_max
            End If

            SetLights
            
            With DAAnims.BodyAnimations(anim_index)
                weapon_index = -1
                If anim_index < DAAnims.NumWeaponAnimations And aa_sk.NumWeapons > 0 Then
                    If WeaponCombo.ListIndex > -1 Then _
                        weapon_index = WeaponCombo.List(WeaponCombo.ListIndex)
                    aux_anim = DAAnims.WeaponAnimations(anim_index).Frames(CurrentFrameScroll.value)
                End If
                
                DrawAASkeleton aa_sk, .Frames(CurrentFrameScroll.value), aux_anim, weapon_index, (DListEnableCheck.value = vbChecked)
                If ShowLastFrameGhostCheck.value = vbChecked And Not aa_sk.IsBattleLocation Then
                    glColorMask GL_TRUE, GL_TRUE, GL_FALSE, GL_TRUE
                    If CurrentFrameScroll.value = 0 Then
                        DrawAASkeleton aa_sk, .Frames(CurrentFrameScroll.max - 1), aux_anim, weapon_index, (DListEnableCheck.value = vbChecked)
                    Else
                        DrawAASkeleton aa_sk, .Frames(CurrentFrameScroll.value - 1), aux_anim, weapon_index, (DListEnableCheck.value = vbChecked)
                    End If
                    glColorMask GL_TRUE, GL_TRUE, GL_TRUE, GL_TRUE
                End If
                SelectAABoneAndModel aa_sk, .Frames(CurrentFrameScroll.value), aux_anim, weapon_index, SelectedBone, SelectedBonePiece
                
                If ShowBonesCheck.value = vbChecked Then
                    glDisable GL_DEPTH_TEST
                    glDisable GL_LIGHTING
                    glColor3f 0, 1, 0
                    DrawAASkeletonBones aa_sk, .Frames(CurrentFrameScroll.value)
                    glEnable GL_DEPTH_TEST
                End If
            End With
    End Select
End Sub
Private Sub ResizePieceX_Change()
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
    
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    Select Case ModelType
        Case K_HRC_SKELETON:
            With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).Model
                .ResizeX = ResizePieceX.value / 100
            End With
        Case K_AA_SKELETON:
            If SelectedBone = aa_sk.NumBones Then
                With aa_sk.WeaponModels(WeaponCombo.List(WeaponCombo.ListIndex))
                    .ResizeX = ResizePieceX.value / 100
                End With
            Else
                With aa_sk.Bones(SelectedBone).Models(SelectedBonePiece)
                    .ResizeX = ResizePieceX.value / 100
                End With
            End If
        Case Else:
            P_Model.ResizeX = ResizePieceX.value / 100
    End Select
    ResizePieceXText.Text = ResizePieceX.value
    Picture1_Paint
    DoNotAddStateQ = False
End Sub
Private Sub ResizePieceY_Change()
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
    
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    Select Case ModelType
        Case K_HRC_SKELETON:
            With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).Model
                .ResizeY = ResizePieceY.value / 100
            End With
        Case K_AA_SKELETON:
            If SelectedBone = aa_sk.NumBones Then
                With aa_sk.WeaponModels(WeaponCombo.List(WeaponCombo.ListIndex))
                    .ResizeY = ResizePieceY.value / 100
                End With
            Else
                With aa_sk.Bones(SelectedBone).Models(SelectedBonePiece)
                    .ResizeY = ResizePieceY.value / 100
                End With
            End If
        Case Else:
            P_Model.ResizeY = ResizePieceY.value / 100
    End Select
    ResizePieceYText.Text = ResizePieceY.value
    Picture1_Paint
    DoNotAddStateQ = False
End Sub
Private Sub ResizePieceZ_Change()
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
        
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    Select Case ModelType
        Case K_HRC_SKELETON:
            With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).Model
                .ResizeZ = ResizePieceZ.value / 100
            End With
        Case K_AA_SKELETON:
            If SelectedBone = aa_sk.NumBones Then
                With aa_sk.WeaponModels(WeaponCombo.List(WeaponCombo.ListIndex))
                    .ResizeZ = ResizePieceZ.value / 100
                End With
            Else
                With aa_sk.Bones(SelectedBone).Models(SelectedBonePiece)
                    .ResizeZ = ResizePieceZ.value / 100
                End With
            End If
        Case Else:
            P_Model.ResizeZ = ResizePieceZ.value / 100
    End Select
    ResizePieceZText.Text = ResizePieceZ.value
    Picture1_Paint
    DoNotAddStateQ = False
End Sub
Private Sub RepositionX_Change()
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
        
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    Select Case ModelType
        Case K_HRC_SKELETON:
            With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).Model
                .RepositionX = RepositionX.value * ComputeDiameter(.BoundingBox) / 100
            End With
        Case K_AA_SKELETON:
            If SelectedBone = aa_sk.NumBones Then
                With aa_sk.WeaponModels(WeaponCombo.List(WeaponCombo.ListIndex))
                    .RepositionX = RepositionX.value * ComputeDiameter(.BoundingBox) / 100
                End With
            Else
                With aa_sk.Bones(SelectedBone).Models(SelectedBonePiece)
                    .RepositionX = RepositionX.value * ComputeDiameter(.BoundingBox) / 100
                End With
            End If
        Case Else:
            P_Model.RepositionX = RepositionX.value * ComputeDiameter(P_Model.BoundingBox) / 100
    End Select
    RepositionXText.Text = RepositionX.value
    Picture1_Paint
    DoNotAddStateQ = False
End Sub
Private Sub RepositionY_Change()
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
        
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    Select Case ModelType
        Case K_HRC_SKELETON:
            With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).Model
                .RepositionY = RepositionY.value * ComputeDiameter(.BoundingBox) / 100
            End With
        Case K_AA_SKELETON:
            If SelectedBone = aa_sk.NumBones Then
                With aa_sk.WeaponModels(WeaponCombo.List(WeaponCombo.ListIndex))
                    .RepositionY = RepositionY.value * ComputeDiameter(.BoundingBox) / 100
                End With
            Else
                With aa_sk.Bones(SelectedBone).Models(SelectedBonePiece)
                    .RepositionY = RepositionY.value * ComputeDiameter(.BoundingBox) / 100
                End With
            End If
        Case Else:
            P_Model.RepositionY = RepositionY.value * ComputeDiameter(P_Model.BoundingBox) / 100
    End Select
    RepositionYText.Text = RepositionY.value
    Picture1_Paint
    DoNotAddStateQ = False
End Sub
Private Sub RepositionZ_Change()
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
        
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    Select Case ModelType
        Case K_HRC_SKELETON:
            With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).Model
                .RepositionZ = RepositionZ.value * ComputeDiameter(.BoundingBox) / 100
            End With
        Case K_AA_SKELETON:
            If SelectedBone = aa_sk.NumBones Then
                With aa_sk.WeaponModels(WeaponCombo.List(WeaponCombo.ListIndex))
                    .RepositionZ = RepositionZ.value * ComputeDiameter(.BoundingBox) / 100
                End With
            Else
                With aa_sk.Bones(SelectedBone).Models(SelectedBonePiece)
                    .RepositionZ = RepositionZ.value * ComputeDiameter(.BoundingBox) / 100
                End With
            End If
        Case Else:
            P_Model.RepositionZ = RepositionZ.value * ComputeDiameter(P_Model.BoundingBox) / 100
    End Select
    RepositionZText.Text = RepositionZ.value
    Picture1_Paint
    DoNotAddStateQ = False
End Sub

Private Sub RotateAlpha_Change()
    PieceRotationModifiersChanged
End Sub

Private Sub RotateBeta_Change()
    PieceRotationModifiersChanged
End Sub

Private Sub RotateGamma_Change()
    PieceRotationModifiersChanged
End Sub
Private Sub RepositionXText_Change()
    If LoadingBonePieceModifiersQ Then _
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
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
        
    If LoadingBonePieceModifiersQ Then _
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
    If LoadingBonePieceModifiersQ Then _
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
    If LoadingBonePieceModifiersQ Then _
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

Private Sub ResizePieceXText_Change()
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
        
    If IsNumeric(ResizePieceXText.Text) Then
        If ResizePieceXText.Text >= 0 And ResizePieceXText.Text <= 400 Then
            ResizePieceX.value = ResizePieceXText.Text
            ResizePieceX_Change
        Else
            ResizePieceXText.Text = 400
            ResizePieceX_Change
        End If
    Else
        Beep
        ResizePieceXText.Text = 0
        ResizePieceX_Change
    End If
End Sub
Private Sub ResizePieceYText_Change()
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
        
    If IsNumeric(ResizePieceYText.Text) Then
        If ResizePieceYText.Text >= 0 And ResizePieceYText.Text <= 400 Then
            ResizePieceY.value = ResizePieceYText.Text
            ResizePieceY_Change
        Else
            ResizePieceYText.Text = 400
            ResizePieceY_Change
        End If
    Else
        Beep
        ResizePieceYText.Text = 0
        ResizePieceY_Change
    End If
End Sub

Private Sub ResizePieceZText_Change()
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
        
    If IsNumeric(ResizePieceZText.Text) Then
        If ResizePieceZText.Text >= 0 And ResizePieceZText.Text <= 400 Then
            ResizePieceZ.value = ResizePieceZText.Text
            ResizePieceZ_Change
        Else
            ResizePieceZText.Text = 400
            ResizePieceZ_Change
        End If
    Else
        Beep
        ResizePieceZText.Text = 0
        ResizePieceZ_Change
    End If
End Sub

Private Sub RepositionZText_Change()
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
        
    If LoadingBonePieceModifiersQ Then _
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
    If LoadingBonePieceModifiersQ Then _
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
Sub SetFrameEditorFields()
    Dim anim_index As Integer
    FrameDataPartOptions.Enabled = True

    LoadingAnimationQ = True
    
    'If FrameDataPartUpDown.value = K_FRAME_ROOT_TRANSLATION Then
        XAnimationFramePartUpDown.max = 999999999
        XAnimationFramePartUpDown.Min = -999999999
        YAnimationFramePartUpDown.max = 999999999
        YAnimationFramePartUpDown.Min = -999999999
        ZAnimationFramePartUpDown.max = 999999999
        ZAnimationFramePartUpDown.Min = -999999999
    'End If

    Select Case FrameDataPartUpDown.value
        Case K_FRAME_BONE_ROTATION
            If SelectedBone > -1 Then
                Select Case ModelType
                    Case K_HRC_SKELETON:
                        With AAnim.Frames(CurrentFrameScroll.value).Rotations(SelectedBone)
                            XAnimationFramePartText.Text = .alpha
                            YAnimationFramePartText.Text = .Beta
                            ZAnimationFramePartText.Text = .Gamma
                            XAnimationFramePartUpDown.value = .alpha * 10000
                            YAnimationFramePartUpDown.value = .Beta * 10000
                            ZAnimationFramePartUpDown.value = .Gamma * 10000
                        End With
                    Case K_AA_SKELETON:
                        If Not aa_sk.IsBattleLocation Then _
                            anim_index = val(BattleAnimationCombo.List(BattleAnimationCombo.ListIndex))
                        If (SelectedBone = aa_sk.NumBones) Then
                            With DAAnims.WeaponAnimations(anim_index).Frames(CurrentFrameScroll.value).Bones(0)
                                XAnimationFramePartText.Text = .alpha
                                YAnimationFramePartText.Text = .Beta
                                ZAnimationFramePartText.Text = .Gamma
                                XAnimationFramePartUpDown.value = .alpha * 10000
                                YAnimationFramePartUpDown.value = .Beta * 10000
                                ZAnimationFramePartUpDown.value = .Gamma * 10000
                            End With
                        Else
                            With DAAnims.BodyAnimations(anim_index).Frames(CurrentFrameScroll.value).Bones(SelectedBone + IIf(aa_sk.NumBones > 1, 1, 0))
                                XAnimationFramePartText.Text = .alpha
                                YAnimationFramePartText.Text = .Beta
                                ZAnimationFramePartText.Text = .Gamma
                                XAnimationFramePartUpDown.value = .alpha * 10000
                                YAnimationFramePartUpDown.value = .Beta * 10000
                                ZAnimationFramePartUpDown.value = .Gamma * 10000
                            End With
                        End If
                End Select
            Else
                XAnimationFramePartText.Text = " "
                YAnimationFramePartText.Text = " "
                ZAnimationFramePartText.Text = " "
                FrameDataPartOptions.Enabled = False
            End If
        Case K_FRAME_ROOT_ROTATION
            Select Case ModelType
                Case K_HRC_SKELETON:
                    With AAnim.Frames(CurrentFrameScroll.value)
                        XAnimationFramePartText.Text = .RootRotationAlpha
                        YAnimationFramePartText.Text = .RootRotationBeta
                        ZAnimationFramePartText.Text = .RootRotationGamma
                        XAnimationFramePartUpDown.value = .RootRotationAlpha * 10000
                        YAnimationFramePartUpDown.value = .RootRotationBeta * 10000
                        ZAnimationFramePartUpDown.value = .RootRotationGamma * 10000
                    End With
                Case K_AA_SKELETON:
                    If Not aa_sk.IsBattleLocation Then _
                            anim_index = val(BattleAnimationCombo.List(BattleAnimationCombo.ListIndex))
                    With DAAnims.BodyAnimations(anim_index).Frames(CurrentFrameScroll.value).Bones(0)
                        XAnimationFramePartText.Text = .alpha
                        YAnimationFramePartText.Text = .Beta
                        ZAnimationFramePartText.Text = .Gamma
                        XAnimationFramePartUpDown.value = .alpha * 10000
                        YAnimationFramePartUpDown.value = .Beta * 10000
                        ZAnimationFramePartUpDown.value = .Gamma * 10000
                    End With
            End Select
        Case K_FRAME_ROOT_TRANSLATION
            Select Case ModelType
                Case K_HRC_SKELETON:
                    With AAnim.Frames(CurrentFrameScroll.value)
                        XAnimationFramePartText.Text = .RootTranslationX
                        YAnimationFramePartText.Text = .RootTranslationY
                        ZAnimationFramePartText.Text = .RootTranslationZ
                        XAnimationFramePartUpDown.value = .RootTranslationX * 10000
                        YAnimationFramePartUpDown.value = .RootTranslationX * 10000
                        ZAnimationFramePartUpDown.value = .RootTranslationX * 10000
                    End With
                Case K_AA_SKELETON:
                    anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
                    If (SelectedBone = aa_sk.NumBones) Then
                        With DAAnims.WeaponAnimations(anim_index).Frames(CurrentFrameScroll.value)
                            XAnimationFramePartText.Text = .X_start
                            YAnimationFramePartText.Text = .Y_start
                            ZAnimationFramePartText.Text = .Z_start
                            XAnimationFramePartUpDown.value = .X_start * 10000
                            YAnimationFramePartUpDown.value = .Y_start * 10000
                            ZAnimationFramePartUpDown.value = .Z_start * 10000
                        End With
                    Else
                        With DAAnims.BodyAnimations(anim_index).Frames(CurrentFrameScroll.value)
                            XAnimationFramePartText.Text = .X_start
                            YAnimationFramePartText.Text = .Y_start
                            ZAnimationFramePartText.Text = .Z_start
                            XAnimationFramePartUpDown.value = .X_start * 10000
                            YAnimationFramePartUpDown.value = .Y_start * 10000
                            ZAnimationFramePartUpDown.value = .Z_start * 10000
                        End With
                    End If
            End Select
    End Select
    
    'If FrameDataPartUpDown.value <> K_FRAME_ROOT_TRANSLATION Then
    '    XAnimationFramePartUpDown.max = 7200000
    '    XAnimationFramePartUpDown.Min = -7200000
    '    YAnimationFramePartUpDown.max = 7200000
    '    YAnimationFramePartUpDown.Min = -7200000
    '    ZAnimationFramePartUpDown.max = 7200000
    '    ZAnimationFramePartUpDown.Min = -7200000
    'End If
    
    LoadingAnimationQ = False
End Sub
Sub SetTextureEditorFields()
    Dim ti As Integer
    
    TextureSelectCombo.Clear
    Select Case ModelType
        Case K_HRC_SKELETON:
            TextureSelectCombo.Clear
            If (SelectedBone > -1 And SelectedBonePiece > -1) Then
                TexturesFrame.Enabled = True
                With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece)
                    For ti = 0 To .NumTextures - 1
                        TextureSelectCombo.AddItem (.textures(ti).tex_file)
                    Next ti
                    If .NumTextures > 0 Then
                        TextureSelectCombo.ListIndex = 0
                    End If
                End With
            Else
                TexturesFrame.Enabled = False
            End If
            Call TextureViewer_Paint
        Case K_AA_SKELETON:
            With aa_sk
                For ti = 0 To .NumTextures - 1
                    TextureSelectCombo.AddItem (ti)
                Next ti
                If .NumTextures > 0 Then
                    TextureSelectCombo.ListIndex = 0
                Else
                    Call TextureViewer_Paint
                End If
            End With
    End Select
End Sub
Sub SetBoneModifiers()
    LoadingBoneModifiersQ = True
    
    Select Case ModelType
        Case K_HRC_SKELETON:
            With hrc_sk.Bones(SelectedBone)
                ResizeBoneXUpDown.value = .ResizeX * 100
                ResizeBoneYUpDown.value = .ResizeY * 100
                ResizeBoneZUpDown.value = .ResizeZ * 100
                
                BoneLengthText.Text = .length
                BoneLengthUpDown.value = .length * 10000
                BoneLengthUpDown.Increment = Abs(BoneLengthUpDown.value / 100)
            End With
            Call SetFrameEditorFields
            
        Case K_AA_SKELETON:
            If (SelectedBone = aa_sk.NumBones) Then
                With aa_sk.WeaponModels(WeaponCombo.ListIndex)
                    ResizeBoneXUpDown.value = .ResizeX * 100
                    ResizeBoneYUpDown.value = .ResizeY * 100
                    ResizeBoneZUpDown.value = .ResizeZ * 100
                    
                    BoneLengthLabel.Visible = False
                    BoneLengthText.Visible = False
                    BoneLengthUpDown.Visible = False
                    AddPieceButton.Visible = False
                    RemovePieceButton.Visible = False
                End With
            Else
                With aa_sk.Bones(SelectedBone)
                    ResizeBoneXUpDown.value = .ResizeX * 100
                    ResizeBoneYUpDown.value = .ResizeY * 100
                    ResizeBoneZUpDown.value = .ResizeZ * 100
                    
                    BoneLengthText.Text = .length
                    BoneLengthUpDown.value = .length * 10000
                    BoneLengthUpDown.Increment = Abs(BoneLengthUpDown.value / 100)
                    
                    BoneLengthText.Visible = True
                    BoneLengthUpDown.Visible = True
                    AddPieceButton.Visible = True
                    RemovePieceButton.Visible = True
                    BoneLengthLabel.Visible = True
                End With
            End If
                
            Call SetFrameEditorFields
    End Select
    
    LoadingBoneModifiersQ = False
End Sub
Sub DestroyCurrentModel()
    Select Case ModelType
        Case K_HRC_SKELETON:
            FreeHRCSkeletonResources hrc_sk
        Case K_AA_SKELETON:
            FreeAASkeletonResources aa_sk
    End Select
End Sub

Sub SetBonePieceModifiers()
    Dim obj As PModel
    Dim Diam As Single
    Dim weapon_index As Integer
    
    LoadingBonePieceModifiersQ = True
    
    Select Case ModelType
        Case K_HRC_SKELETON:
            obj = hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).Model
            Diam = ComputeDiameter(obj.BoundingBox) / 100
        Case K_AA_SKELETON:
            If SelectedBone = aa_sk.NumBones Then
                weapon_index = WeaponCombo.List(WeaponCombo.ListIndex)
                obj = aa_sk.WeaponModels(weapon_index)
                Diam = ComputeDiameter(obj.BoundingBox) / 100
            Else
                obj = aa_sk.Bones(SelectedBone).Models(SelectedBonePiece)
                Diam = ComputeDiameter(obj.BoundingBox) / 100
            End If
        Case Else:
            obj = P_Model
            Diam = ComputeDiameter(P_Model.BoundingBox) / 100
    End Select
    
    With obj
        ResizePieceX.value = .ResizeX * 100
        ResizePieceY.value = .ResizeY * 100
        ResizePieceZ.value = .ResizeZ * 100
        
        RepositionX.value = .RepositionX / Diam
        RepositionY.value = .RepositionY / Diam
        RepositionZ.value = .RepositionZ / Diam
        
        RotateAlpha.value = .RotateAlpha
        RotateBeta.value = .RotateBeta
        RotateGamma.value = .RotateGamma
        
        ResizePieceXText.Text = .ResizeX * 100
        ResizePieceYText.Text = .ResizeY * 100
        ResizePieceZText.Text = .ResizeZ * 100
        
        RepositionXText.Text = .RepositionX / Diam
        RepositionYText.Text = .RepositionY / Diam
        RepositionZText.Text = .RepositionZ / Diam
        
        RotateAlphaText.Text = .RotateAlpha
        RotateBetaText.Text = .RotateBeta
        RotateGammaText.Text = .RotateGamma
        
        ResizePieceX.Refresh
        ResizePieceY.Refresh
        ResizePieceZ.Refresh
        
        RepositionX.Refresh
        RepositionY.Refresh
        RepositionZ.Refresh
        
        RotateAlpha.Refresh
        RotateBeta.Refresh
        RotateGamma.Refresh
    End With
    
    LoadingBonePieceModifiersQ = False
End Sub
Sub UpdateEditedPiece()
    AddStateToBuffer
    Select Case ModelType
        Case K_HRC_SKELETON:
            hrc_sk.Bones(EditedBone).Resources(EditedBonePiece).Model = EditedPModel
            SetOGLContext Picture1.hdc, OGLContext
            CreateDListsFromHRCSkeleton hrc_sk
        Case K_AA_SKELETON:
            SetOGLContext Picture1.hdc, OGLContext
            If EditedBone = aa_sk.NumBones Then
                aa_sk.WeaponModels(WeaponCombo.List(WeaponCombo.ListIndex)) = EditedPModel
                CreateDListsFromPModel aa_sk.WeaponModels(WeaponCombo.List(WeaponCombo.ListIndex))
            Else
                aa_sk.Bones(EditedBone).Models(EditedBonePiece) = EditedPModel
                CreateDListsFromAASkeleton aa_sk
            End If
            
        Case Else:
            P_Model = EditedPModel
            SetOGLContext Picture1.hdc, OGLContext
            CreateDListsFromPModel EditedPModel
    End Select
    SelectedBone = EditedBone
    SelectedBonePiece = EditedBonePiece
    SetBonePieceModifiers
End Sub
Sub FillBoneSelector()
    Dim BI
    
    BoneSelector.Clear
    
    Select Case ModelType
        Case K_HRC_SKELETON:
            With hrc_sk
                For BI = 0 To .NumBones - 1
                    BoneSelector.AddItem (.Bones(BI).joint_i + "-" + .Bones(BI).joint_f)
                Next BI
            End With
            BoneSelector.Enabled = True
        Case K_AA_SKELETON:
            With aa_sk
                For BI = 0 To .NumBones - 1
                    BoneSelector.AddItem ("Joint" + Str$(.Bones(BI).ParentBone) + "- Joint" + Str$(BI))
                Next BI
                If .NumWeapons > 0 And .NumWeaponAnims > 0 Then _
                    BoneSelector.AddItem ("Weapon")
            End With
            BoneSelector.Enabled = True
        Case Else:
            BoneSelector.Enabled = False
    End Select
End Sub
Sub SetLights()
    Dim light_x As Single
    Dim light_y As Single
    Dim light_z As Single
    
    Dim p_min As Point3D
    Dim p_max As Point3D
    
    Dim scene_diameter As Single
    
    Dim anim_index As Integer
    
    Dim infinityFarQ As Boolean
    
    If FrontLight.value = vbUnchecked And _
       RearLight.value = vbUnchecked And _
       RightLight.value = vbUnchecked And _
       LeftLight.value = vbUnchecked Then
        glDisable GL_LIGHTING
        Exit Sub
    End If
    
    Select Case ModelType
        Case K_P_FIELD_MODEL:
            ComputePModelBoundingBox P_Model, p_min, p_max
        Case K_P_BATTLE_MODEL:
            ComputePModelBoundingBox P_Model, p_min, p_max
        Case K_HRC_SKELETON:
            ComputeHRCBoundingBox hrc_sk, AAnim.Frames(CurrentFrameScroll.value), p_min, p_max
        Case K_AA_SKELETON:
            If Not aa_sk.IsBattleLocation Then _
                anim_index = val(BattleAnimationCombo.List(BattleAnimationCombo.ListIndex))
            ComputeAABoundingBox aa_sk, _
                DAAnims.BodyAnimations(anim_index).Frames(CurrentFrameScroll.value), _
                p_min, p_max
    End Select
    
    scene_diameter = -2 * ComputeSceneRadius(p_min, p_max)
    
    light_x = scene_diameter / LIGHT_STEPS * LightPosXScroll.value
    light_y = scene_diameter / LIGHT_STEPS * LightPosYScroll.value
    light_z = scene_diameter / LIGHT_STEPS * LightPosZScroll.value
    
    infinityFarQ = (InifintyFarLightsCheck.value = vbChecked)
    
    If RightLight.value = vbChecked Then
        SetLighting GL_LIGHT0, light_z, light_y, light_x, 0.5, 0.5, 0.5, infinityFarQ
    Else
        glDisable GL_LIGHT0
    End If
    
    If LeftLight.value = vbChecked Then
        SetLighting GL_LIGHT1, -light_z, light_y, light_x, 0.5, 0.5, 0.5, infinityFarQ
    Else
        glDisable GL_LIGHT1
    End If
    
    If FrontLight.value = vbChecked Then
        SetLighting GL_LIGHT2, light_x, light_y, light_z, 1, 1, 1, infinityFarQ
    Else
        glDisable GL_LIGHT2
    End If
    
    If RearLight.value = vbChecked Then
        SetLighting GL_LIGHT3, light_x, light_y, -light_z, 0.75, 0.75, 0.75, infinityFarQ
    Else
        glDisable GL_LIGHT3
    End If
End Sub

Private Sub XAnimationFramePartUpDown_Change()
    Dim anim_index As Integer
    Dim frame_index As Integer
    Dim num_frames As Integer
    Dim fi As Integer
    Dim val As Single
    Dim diff As Single
    
    If LoadingAnimationQ Then _
        Exit Sub
        
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    
    val = XAnimationFramePartUpDown.value / 10000
    XAnimationFramePartText.Text = val

    frame_index = CurrentFrameScroll.value
    num_frames = frame_index
    'Must propagate the changes to the following frames?
    If PropagateChangesForwardCheck.value = vbChecked Then _
        num_frames = CurrentFrameScroll.max - 1
    Select Case FrameDataPartUpDown.value
        Case K_FRAME_BONE_ROTATION
            If SelectedBone > -1 Then
                Select Case ModelType
                    Case K_HRC_SKELETON:
                        diff = val - AAnim.Frames(frame_index).Rotations(SelectedBone).alpha
                        For fi = frame_index To num_frames
                            With AAnim.Frames(fi).Rotations(SelectedBone)
                                .alpha = .alpha + diff
                            End With
                        Next fi
                    Case K_AA_SKELETON:
                        anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
                        If (SelectedBone = aa_sk.NumBones) Then
                            diff = val - _
                                DAAnims.WeaponAnimations(anim_index).Frames(frame_index).Bones(0).alpha
                            For fi = frame_index To num_frames
                                With DAAnims.WeaponAnimations(anim_index).Frames(fi).Bones(0)
                                    .alpha = .alpha + diff
                                End With
                            Next fi
                        Else
                            diff = val - _
                                DAAnims.BodyAnimations(anim_index).Frames(frame_index).Bones(SelectedBone + 1).alpha
                            For fi = frame_index To num_frames
                                With DAAnims.BodyAnimations(anim_index).Frames(fi).Bones(SelectedBone + 1)
                                    .alpha = .alpha + diff
                                End With
                            Next fi
                        End If
                End Select
            End If
        Case K_FRAME_ROOT_ROTATION
            Select Case ModelType
                Case K_HRC_SKELETON:
                    diff = val - _
                        AAnim.Frames(frame_index).RootRotationAlpha
                    For fi = frame_index To num_frames
                        With AAnim.Frames(fi)
                            .RootRotationAlpha = .RootRotationAlpha + diff
                        End With
                    Next fi
                Case K_AA_SKELETON:
                    anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
                    diff = val - _
                        DAAnims.BodyAnimations(anim_index).Frames(frame_index).Bones(0).alpha
                    For fi = frame_index To num_frames
                        With DAAnims.BodyAnimations(anim_index).Frames(fi).Bones(0)
                            .alpha = .alpha + diff
                        End With
                    Next fi
            End Select
        Case K_FRAME_ROOT_TRANSLATION
            Select Case ModelType
                Case K_HRC_SKELETON:
                    diff = val - _
                        AAnim.Frames(frame_index).RootTranslationX
                    For fi = frame_index To num_frames
                        With AAnim.Frames(fi)
                            .RootTranslationX = .RootTranslationX + diff
                        End With
                    Next fi
                Case K_AA_SKELETON:
                    anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
                    If (SelectedBone = aa_sk.NumBones) Then
                        diff = val - _
                            DAAnims.WeaponAnimations(anim_index).Frames(frame_index).X_start
                        For fi = frame_index To num_frames
                            With DAAnims.WeaponAnimations(anim_index).Frames(fi)
                                .X_start = .X_start + diff
                            End With
                        Next fi
                    Else
                        diff = val - _
                            DAAnims.BodyAnimations(anim_index).Frames(frame_index).X_start
                        For fi = frame_index To num_frames
                            With DAAnims.BodyAnimations(anim_index).Frames(fi)
                                .X_start = .X_start + diff
                            End With
                        Next fi
                        If aa_sk.NumWeapons > 0 Then
                            For fi = frame_index To num_frames
                                With DAAnims.WeaponAnimations(anim_index).Frames(fi)
                                    .X_start = .X_start + diff
                                End With
                            Next fi
                        End If
                    End If
            End Select
    End Select
    Picture1_Paint
    DoNotAddStateQ = False
End Sub

Private Sub YAnimationFramePartUpDown_Change()
    Dim anim_index As Integer
    Dim frame_index As Integer
    Dim num_frames As Integer
    Dim fi As Integer
    Dim val As Single
    Dim diff As Single
    
    If LoadingAnimationQ Then _
        Exit Sub
        
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    
    val = YAnimationFramePartUpDown.value / 10000
    YAnimationFramePartText.Text = val

    frame_index = CurrentFrameScroll.value
    num_frames = frame_index
    'Must propagate the changes to the following frames?
    If PropagateChangesForwardCheck.value = vbChecked Then _
        num_frames = CurrentFrameScroll.max - 1
    Select Case FrameDataPartUpDown.value
        Case K_FRAME_BONE_ROTATION
            If SelectedBone > -1 Then
                Select Case ModelType
                    Case K_HRC_SKELETON:
                        diff = val - AAnim.Frames(frame_index).Rotations(SelectedBone).Beta
                        For fi = frame_index To num_frames
                            With AAnim.Frames(fi).Rotations(SelectedBone)
                                .Beta = .Beta + diff
                            End With
                        Next fi
                    Case K_AA_SKELETON:
                        anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
                        If (SelectedBone = aa_sk.NumBones) Then
                            diff = val - _
                                DAAnims.WeaponAnimations(anim_index).Frames(frame_index).Bones(0).Beta
                            For fi = frame_index To num_frames
                                With DAAnims.WeaponAnimations(anim_index).Frames(fi).Bones(0)
                                    .Beta = .Beta + diff
                                End With
                            Next fi
                        Else
                            diff = val - _
                                DAAnims.BodyAnimations(anim_index).Frames(frame_index).Bones(SelectedBone + 1).Beta
                            For fi = frame_index To num_frames
                                With DAAnims.BodyAnimations(anim_index).Frames(fi).Bones(SelectedBone + 1)
                                    .Beta = .Beta + diff
                                End With
                            Next fi
                        End If
                End Select
            End If
        Case K_FRAME_ROOT_ROTATION
            Select Case ModelType
                Case K_HRC_SKELETON:
                    diff = val - _
                        AAnim.Frames(frame_index).RootRotationBeta
                    For fi = frame_index To num_frames
                        With AAnim.Frames(fi)
                            .RootRotationBeta = .RootRotationBeta + diff
                        End With
                    Next fi
                Case K_AA_SKELETON:
                    anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
                    diff = val - _
                        DAAnims.BodyAnimations(anim_index).Frames(frame_index).Bones(0).Beta
                    For fi = frame_index To num_frames
                        With DAAnims.BodyAnimations(anim_index).Frames(fi).Bones(0)
                            .Beta = .Beta + diff
                        End With
                    Next fi
            End Select
        Case K_FRAME_ROOT_TRANSLATION
            Select Case ModelType
                Case K_HRC_SKELETON:
                    diff = val - _
                        AAnim.Frames(frame_index).RootTranslationY
                    For fi = frame_index To num_frames
                        With AAnim.Frames(fi)
                            .RootTranslationY = .RootTranslationY + diff
                        End With
                    Next fi
                Case K_AA_SKELETON:
                    anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
                    If (SelectedBone = aa_sk.NumBones) Then
                        diff = val - _
                            DAAnims.WeaponAnimations(anim_index).Frames(frame_index).Y_start
                        For fi = frame_index To num_frames
                            With DAAnims.WeaponAnimations(anim_index).Frames(fi)
                                .Y_start = .Y_start + diff
                            End With
                        Next fi
                    Else
                        diff = val - _
                            DAAnims.BodyAnimations(anim_index).Frames(frame_index).Y_start
                        For fi = frame_index To num_frames
                            With DAAnims.BodyAnimations(anim_index).Frames(fi)
                                .Y_start = .Y_start + diff
                            End With
                        Next fi
                        If aa_sk.NumWeapons > 0 Then
                            For fi = frame_index To num_frames
                                With DAAnims.WeaponAnimations(anim_index).Frames(fi)
                                    .Y_start = .Y_start + diff
                                End With
                            Next fi
                        End If
                    End If
            End Select
    End Select
    Picture1_Paint
    DoNotAddStateQ = False
End Sub

Private Sub ZAnimationFramePartUpDown_Change()
    Dim anim_index As Integer
    Dim frame_index As Integer
    Dim num_frames As Integer
    Dim fi As Integer
    Dim val As Single
    Dim diff As Single
    
    If LoadingAnimationQ Then _
        Exit Sub
        
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    
    val = ZAnimationFramePartUpDown.value / 10000
    ZAnimationFramePartText.Text = val

    frame_index = CurrentFrameScroll.value
    num_frames = frame_index
    'Must propagate the changes to the following frames?
    If PropagateChangesForwardCheck.value = vbChecked Then _
        num_frames = CurrentFrameScroll.max - 1
    Select Case FrameDataPartUpDown.value
        Case K_FRAME_BONE_ROTATION
            If SelectedBone > -1 Then
                Select Case ModelType
                    Case K_HRC_SKELETON:
                        diff = val - AAnim.Frames(frame_index).Rotations(SelectedBone).Gamma
                        For fi = frame_index To num_frames
                            With AAnim.Frames(fi).Rotations(SelectedBone)
                                .Gamma = .Gamma + diff
                            End With
                        Next fi
                    Case K_AA_SKELETON:
                        anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
                        If (SelectedBone = aa_sk.NumBones) Then
                            diff = val - _
                                DAAnims.WeaponAnimations(anim_index).Frames(frame_index).Bones(0).Gamma
                            For fi = frame_index To num_frames
                                With DAAnims.WeaponAnimations(anim_index).Frames(fi).Bones(0)
                                    .Gamma = .Gamma + diff
                                End With
                            Next fi
                        Else
                            diff = val - _
                                DAAnims.BodyAnimations(anim_index).Frames(frame_index).Bones(SelectedBone + 1).Gamma
                            For fi = frame_index To num_frames
                                With DAAnims.BodyAnimations(anim_index).Frames(fi).Bones(SelectedBone + 1)
                                    .Gamma = .Gamma + diff
                                End With
                            Next fi
                        End If
                End Select
            End If
        Case K_FRAME_ROOT_ROTATION
            Select Case ModelType
                Case K_HRC_SKELETON:
                    diff = val - _
                        AAnim.Frames(frame_index).RootRotationGamma
                    For fi = frame_index To num_frames
                        With AAnim.Frames(fi)
                            .RootRotationGamma = .RootRotationGamma + diff
                        End With
                    Next fi
                Case K_AA_SKELETON:
                    anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
                    diff = val - _
                        DAAnims.BodyAnimations(anim_index).Frames(frame_index).Bones(0).Gamma
                    For fi = frame_index To num_frames
                        With DAAnims.BodyAnimations(anim_index).Frames(frame_index).Bones(0)
                            .Gamma = .Gamma + diff
                        End With
                    Next fi
            End Select
        Case K_FRAME_ROOT_TRANSLATION
            Select Case ModelType
                Case K_HRC_SKELETON:
                    diff = val - _
                        AAnim.Frames(frame_index).RootTranslationZ
                    For fi = frame_index To num_frames
                        With AAnim.Frames(fi)
                            .RootTranslationZ = .RootTranslationZ + diff
                        End With
                    Next fi
                Case K_AA_SKELETON:
                    anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
                    If (SelectedBone = aa_sk.NumBones) Then
                        diff = val - _
                            DAAnims.WeaponAnimations(anim_index).Frames(frame_index).Z_start
                        For fi = frame_index To num_frames
                            With DAAnims.WeaponAnimations(anim_index).Frames(fi)
                                .Z_start = .Z_start + diff
                            End With
                        Next fi
                    Else
                        diff = val - _
                            DAAnims.BodyAnimations(anim_index).Frames(frame_index).Z_start
                        For fi = frame_index To num_frames
                            With DAAnims.BodyAnimations(anim_index).Frames(fi)
                                .Z_start = .Z_start + diff
                            End With
                        Next fi
                        If aa_sk.NumWeapons > 0 Then
                            For fi = frame_index To num_frames
                                With DAAnims.WeaponAnimations(anim_index).Frames(fi)
                                    .Z_start = .Z_start + diff
                                End With
                            Next fi
                        End If
                    End If
            End Select
    End Select
    Picture1_Paint
    DoNotAddStateQ = False
End Sub

Private Sub WeaponCombo_Click()
    Picture1_Paint
End Sub

Private Sub ZeroAsTransparent_Click()
    Dim newZeroTransparentValue As Integer
    newZeroTransparentValue = IIf(ZeroAsTransparent.value = vbChecked, 1, 0)
    Select Case ModelType
        Case K_HRC_SKELETON:
            If SelectedBone > -1 And SelectedBonePiece > -1 Then
                With hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece)
                    If (.textures(TextureSelectCombo.ListIndex).tex_id <> -1) Then
                        If (.textures(TextureSelectCombo.ListIndex).ColorKeyFlag _
                            <> newZeroTransparentValue) Then
                                UnloadTexture .textures(TextureSelectCombo.ListIndex)
                                LoadTEXTexture .textures(TextureSelectCombo.ListIndex)
                                LoadBitmapFromTEXTexture .textures(TextureSelectCombo.ListIndex)
                        End If
                    End If
                End With
            End If
        Case K_AA_SKELETON:
            With aa_sk
                If (.textures(TextureSelectCombo.ListIndex).tex_id <> -1) Then
                    If (.textures(TextureSelectCombo.ListIndex).ColorKeyFlag _
                        <> newZeroTransparentValue) Then
                            UnloadTexture .textures(TextureSelectCombo.ListIndex)
                            LoadTEXTexture .textures(TextureSelectCombo.ListIndex)
                            LoadBitmapFromTEXTexture .textures(TextureSelectCombo.ListIndex)
                    End If
                End If
            End With
    End Select
    Picture1_Paint
End Sub
Private Sub DrawShadow(ByRef p_min As Point3D, ByRef p_max As Point3D)
    Dim ground_radius As Single
    Dim num_segments As Integer
    Dim sub_y As Single
    Dim p_min_aux As Point3D
    Dim p_max_aux As Point3D
    Dim anim_index As Integer
    Dim si As Integer
    Dim PI As Double
    PI = 3.141593
    
    Dim cx As Single
    Dim CZ As Single
    
    cx = 0
    CZ = 0
    
    sub_y = p_max.y
    p_min_aux = p_min
    p_max_aux = p_max
    p_min_aux.y = 0
    p_max_aux.y = 0
    cx = (p_min.x + p_max.x) / 2
    CZ = (p_min.z + p_max.z) / 2
    ground_radius = CalculateDistance(p_min_aux, p_max_aux) / 2

    'Draw shadow
    num_segments = 20
    glBegin GL_TRIANGLE_FAN
        glColor4f 0.1, 0.1, 0.1, 0.5
        glVertex3f cx, 0, CZ
        For si = 0 To num_segments
            glColor4f 0.1, 0.1, 0.1, 0
            glVertex3f ground_radius * Sin(si * 2 * PI / num_segments) + cx, 0, _
                        ground_radius * Cos(si * 2 * PI / num_segments) + CZ
        Next si
    glEnd
    glEnable GL_DEPTH_TEST
    glDisable GL_FOG
    
    'Draw underlying box (just depth)
    glColorMask GL_FALSE, GL_FALSE, GL_FALSE, GL_FALSE
    glColor3f 1, 1, 1
    glBegin GL_QUADS
        glVertex3f p_max.x, 0, p_max.z
        glVertex3f p_max.x, 0, p_min.z
        glVertex3f p_min.x, 0, p_min.z
        glVertex3f p_min.x, 0, p_max.z
    
        glVertex3f p_max.x, 0, p_max.z
        glVertex3f p_max.x, sub_y, p_max.z
        glVertex3f p_max.x, sub_y, p_min.z
        glVertex3f p_max.x, 0, p_min.z
        
        glVertex3f p_max.x, 0, p_min.z
        glVertex3f p_max.x, sub_y, p_min.z
        glVertex3f p_min.x, sub_y, p_min.z
        glVertex3f p_min.x, 0, p_min.z
        
        glVertex3f p_min.x, sub_y, p_max.z
        glVertex3f p_min.x, 0, p_max.z
        glVertex3f p_min.x, 0, p_min.z
        glVertex3f p_min.x, sub_y, p_min.z
        
        glVertex3f p_max.x, sub_y, p_max.z
        glVertex3f p_max.x, 0, p_max.z
        glVertex3f p_min.x, 0, p_max.z
        glVertex3f p_min.x, sub_y, p_max.z
    glEnd
    glColorMask GL_TRUE, GL_TRUE, GL_TRUE, GL_TRUE
    
End Sub

Private Sub DrawGround()
    Dim gi As Integer
    Dim r As Single
    Dim g As Single
    Dim B As Single
    Dim lw As Integer
    
    glMatrixMode GL_PROJECTION
    glPushMatrix
    glLoadIdentity
    gluPerspective 60, Picture1.ScaleWidth / Picture1.ScaleHeight, 0.1, 1000000
    
    Dim f_mode As Long
    Dim f_color(4) As Single
    Dim f_end As Single
    Dim f_start As Single
    
    f_color(0) = 0.5
    f_color(1) = 0.5
    f_color(2) = 1
    f_mode = GL_LINEAR
    glEnable GL_FOG
    glFogiv GL_FOG_MODE, f_mode
    glFogfv GL_FOG_COLOR, f_color(0)
    
    
    f_end = 100000
    f_start = 10000
    glFogfv GL_FOG_END, f_end
    glFogfv GL_FOG_START, f_start
    
    'Draw plane
    glColor3f 0.9, 0.9, 1
    glDisable GL_DEPTH_TEST
    glBegin GL_QUADS
        glVertex4f 1, 0, 1, 0.000001
        glVertex4f 1, 0, -1, 0.000001
        glVertex4f -1, 0, -1, 0.000001
        glVertex4f -1, 0, 1, 0.000001
    glEnd
    
    r = 0.9
    g = 0.9
    B = 1
    lw = 10
    'glEnable (GL_LINE_SMOOTH)
    For gi = 10 To 5 Step -1
        glLineWidth lw
        glColor3f r, g, B
        glBegin GL_LINES
            glVertex4f 0#, 0#, 1#, 0.000001
            glVertex4f 0#, 0#, -1#, 0.000001
            glVertex4f -1#, 0#, 0#, 0.000001
            glVertex4f 1#, 0#, 0#, 0.000001
        glEnd
        r = 0.9 - 0.9 / 10# * (10 - gi)
        g = 0.9 - 0.9 / 10# * (10 - gi)
        B = 1 - 1# / 10# * (10 - gi)
        lw = lw - 2
    Next gi
    glLineWidth 1
    'glDisable (GL_LINE_SMOOTH)
    
    glEnable GL_DEPTH_TEST
    glDisable GL_FOG
    glClear GL_DEPTH_BUFFER_BIT
    glPopMatrix
End Sub
Sub SetWeaponAnimationAttachedToBone(ByVal middleQ As Boolean)
    Dim anim_index As Integer
    Dim fi As Integer
    Dim MV_matrix(16) As Double
    Dim jsp As Integer
    
    AddStateToBuffer
    anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
    With DAAnims.BodyAnimations(anim_index)
        glMatrixMode GL_MODELVIEW
        glPushMatrix
        glLoadIdentity
        For fi = 0 To .NumFrames2 - 1
            If middleQ Then
                jsp = MoveToAABoneMiddle(aa_sk, DAAnims.BodyAnimations(anim_index).Frames(fi), SelectedBone)
            Else
                jsp = MoveToAABoneEnd(aa_sk, DAAnims.BodyAnimations(anim_index).Frames(fi), SelectedBone)
            End If
            glGetDoublev GL_MODELVIEW_MATRIX, MV_matrix(0)
            DAAnims.WeaponAnimations(anim_index).Frames(fi).X_start = MV_matrix(12)
            DAAnims.WeaponAnimations(anim_index).Frames(fi).Y_start = MV_matrix(13)
            DAAnims.WeaponAnimations(anim_index).Frames(fi).Z_start = MV_matrix(14)
            While jsp > 0
                glPopMatrix
                jsp = jsp - 1
            Wend
        Next fi
        glPopMatrix
    End With
    SelectBoneForWeaponAttachmentQ = False
    SetFrameEditorFields
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
    
    If Editor.Visible = True Then
        MsgBox "Can't ReDo while the editor window is open"
        Exit Sub
    End If

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
    
    If Editor.Visible = True Then
        MsgBox "Can't UnDo while the editor window is open"
        Exit Sub
    End If

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

Private Sub RestoreState(ByRef State As ModelEditorState)
    With State
        SelectedBone = .SelectedBone
        SelectedBonePiece = .SelectedBonePiece
        
        alpha = .alpha
        Beta = .Beta
        Gamma = .Gamma
        
        DIST = .DIST
        
        PanX = .PanX
        PanY = .PanY
        PanZ = .PanZ
    
        Select Case (ModelType)
            Case K_HRC_SKELETON:
                hrc_sk = .hrc_sk
                AAnim = .AAnim
                
                If (SelectedBone > -1) Then
                    SetBoneModifiers
                    If (SelectedBonePiece > -1) Then
                        SetBonePieceModifiers
                        SetTextureEditorFields
                    End If
                End If
                SetTextureEditorFields
                SetFrameEditorFields
                
                TextureSelectCombo.ListIndex = .TextureIndex
                If .FrameIndex > AAnim.NumFrames Then
                    CurrentFrameScroll.value = .FrameIndex
                    CurrentFrameScroll.max = AAnim.NumFrames
                Else
                    CurrentFrameScroll.max = AAnim.NumFrames
                    CurrentFrameScroll.value = .FrameIndex
                End If
            Case K_AA_SKELETON:
                aa_sk = .aa_sk
                DAAnims = .DAAnims
                
                If (SelectedBone > -1) Then
                    SetBoneModifiers
                    If (SelectedBonePiece > -1) Then _
                        SetBonePieceModifiers
                End If
                SetTextureEditorFields
                SetFrameEditorFields
                
                If BattleAnimationCombo.Visible Then _
                    BattleAnimationCombo.ListIndex = .BattleAnimationIndex
                    If .FrameIndex < .DAAnims.BodyAnimations(.BattleAnimationIndex).NumFrames2 Then
                        CurrentFrameScroll.value = .FrameIndex
                        CurrentFrameScroll.max = .DAAnims.BodyAnimations(.BattleAnimationIndex).NumFrames2
                    Else
                        CurrentFrameScroll.max = .DAAnims.BodyAnimations(.BattleAnimationIndex).NumFrames2
                        CurrentFrameScroll.value = .FrameIndex
                    End If
                If WeaponCombo.Visible Then _
                    WeaponCombo.ListIndex = .WeaponIndex
                TextureSelectCombo.ListIndex = .TextureIndex
            Case K_P_BATTLE_MODEL:
                P_Model = .P_Model
            Case K_P_FIELD_MODEL:
                P_Model = .P_Model
        End Select
    End With
End Sub

Private Sub StoreState(ByRef State As ModelEditorState)
    With State
        .SelectedBone = SelectedBone
        .SelectedBonePiece = SelectedBonePiece
        
        .alpha = alpha
        .Beta = Beta
        .Gamma = Gamma
        
        .DIST = DIST
        
        .PanX = PanX
        .PanY = PanY
        .PanZ = PanZ
    
        Select Case (ModelType)
            Case K_HRC_SKELETON:
                .hrc_sk = hrc_sk
                .AAnim = AAnim
                
                .FrameIndex = CurrentFrameScroll.value
                .TextureIndex = TextureSelectCombo.ListIndex
            Case K_AA_SKELETON:
                .aa_sk = aa_sk
                .DAAnims = DAAnims
                
                .FrameIndex = CurrentFrameScroll.value
                If BattleAnimationCombo.Visible Then _
                    .BattleAnimationIndex = BattleAnimationCombo.ListIndex
                If WeaponCombo.Visible Then _
                    .WeaponIndex = WeaponCombo.ListIndex
                .TextureIndex = TextureSelectCombo.ListIndex
            Case K_P_BATTLE_MODEL:
                .P_Model = P_Model
            Case K_P_FIELD_MODEL:
                .P_Model = P_Model
        End Select
    End With
End Sub
Private Sub ResetCamera()
    Dim p_min As Point3D
    Dim p_max As Point3D
    
    Dim anim_index As Integer
    
    If loaded Then
        Select Case (ModelType)
            Case K_HRC_SKELETON:
                ComputeHRCBoundingBox hrc_sk, AAnim.Frames(CurrentFrameScroll.value), p_min, p_max
            Case K_AA_SKELETON:
                If Not aa_sk.IsBattleLocation Then _
                    anim_index = BattleAnimationCombo.List(BattleAnimationCombo.ListIndex)
                ComputeAABoundingBox aa_sk, _
                    DAAnims.BodyAnimations(anim_index).Frames(CurrentFrameScroll.value), _
                    p_min, p_max
            Case K_P_BATTLE_MODEL:
                ComputePModelBoundingBox P_Model, p_min, p_max
            Case K_P_FIELD_MODEL:
                ComputePModelBoundingBox P_Model, p_min, p_max
        End Select
        
        alpha = 200
        Beta = 45
        Gamma = 0
        PanX = 0
        PanY = 0
        PanZ = 0
        DIST = -2 * ComputeSceneRadius(p_min, p_max)
    End If
End Sub
Private Sub SetOGLSettings()
    SetOGLContext Picture1.hdc, OGLContext
        
    glClearDepth 1#
    glDepthFunc GL_LEQUAL
    glEnable GL_DEPTH_TEST
    glEnable GL_BLEND
    glEnable GL_ALPHA_TEST
    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
    glAlphaFunc GL_GREATER, 0
    glCullFace GL_FRONT
    glEnable GL_CULL_FACE
    
    'init the function
    'InitAddrCall
    
    'get your api entries
    'Initlpfns
    
    Call Picture1_Paint
End Sub

Private Sub PieceRotationModifiersChanged()
    If LoadingBonePieceModifiersQ Then _
        Exit Sub
        
    If Not DoNotAddStateQ Then AddStateToBuffer
    DoNotAddStateQ = True
    Select Case ModelType
        Case K_HRC_SKELETON:
            RotatePModelModifiers hrc_sk.Bones(SelectedBone).Resources(SelectedBonePiece).Model, _
                    RotateAlpha.value, RotateBeta.value, RotateGamma.value
        Case K_AA_SKELETON:
            If SelectedBone = aa_sk.NumBones Then
                RotatePModelModifiers aa_sk.WeaponModels(WeaponCombo.List(WeaponCombo.ListIndex)), _
                    RotateAlpha.value, RotateBeta.value, RotateGamma.value
            Else
                RotatePModelModifiers aa_sk.Bones(SelectedBone).Models(SelectedBonePiece), _
                    RotateAlpha.value, RotateBeta.value, RotateGamma.value
            End If
        Case Else:
            RotatePModelModifiers P_Model, RotateAlpha.value, RotateBeta.value, RotateGamma.value
    End Select
    RotateAlphaText.Text = RotateAlpha.value
    RotateBetaText.Text = RotateBeta.value
    RotateGammaText.Text = RotateGamma.value
    Picture1_Paint
    DoNotAddStateQ = False
End Sub

