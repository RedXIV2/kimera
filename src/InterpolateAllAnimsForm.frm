VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form InterpolateAllAnimsForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interpolate all FF7 animations"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame ProgressFrame 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Width           =   6855
      Begin VB.PictureBox ProgressBarPicture 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   357
         TabIndex        =   16
         Top             =   360
         Width           =   5415
      End
      Begin VB.CommandButton CancelCommand 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5640
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label ProgressLabel 
         AutoSize        =   -1  'True
         Caption         =   "Progress"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Frame InterpolateOptionsFrame 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton SaveConfigCommand 
         Caption         =   "Save config"
         Height          =   375
         Left            =   5400
         TabIndex        =   32
         Top             =   1680
         Width           =   1335
      End
      Begin MSComCtl2.UpDown NumInterpFramesBattleUpDown 
         Height          =   285
         Left            =   5760
         TabIndex        =   30
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   3
         BuddyControl    =   "NumInterpFramesBattleText"
         BuddyDispid     =   196638
         OrigLeft        =   5760
         OrigTop         =   960
         OrigRight       =   6015
         OrigBottom      =   1245
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox NumInterpFramesBattleText 
         Height          =   285
         Left            =   5400
         TabIndex        =   29
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox NumInterpFramesFieldText 
         Height          =   285
         Left            =   5400
         TabIndex        =   28
         Top             =   480
         Width           =   375
      End
      Begin MSComCtl2.UpDown NumInterpFramesFieldUpDown 
         Height          =   285
         Left            =   5760
         TabIndex        =   27
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "NumInterpFramesFieldText"
         BuddyDispid     =   196636
         OrigLeft        =   5760
         OrigTop         =   480
         OrigRight       =   6015
         OrigBottom      =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton MagicLGPDataDirDestCommand 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   25
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton BattleLGPDataDirDestCommand 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   24
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton CharLGPDataDirDestCommand 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox MagicLGPDataDirDestText 
         Height          =   285
         Left            =   3120
         TabIndex        =   21
         Top             =   1200
         Width           =   1600
      End
      Begin VB.TextBox BattleLGPDataDirDestText 
         Height          =   285
         Left            =   3120
         TabIndex        =   20
         Top             =   840
         Width           =   1600
      End
      Begin VB.TextBox CharLGPDataDirDestText 
         Height          =   285
         Left            =   3120
         TabIndex        =   19
         Top             =   480
         Width           =   1600
      End
      Begin VB.CommandButton GoCommand 
         Caption         =   "Go!"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   5175
      End
      Begin VB.CheckBox MagicLGPDataDirCheck 
         Height          =   255
         Left            =   6360
         TabIndex        =   12
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox BattleLGPDataDirCheck 
         Height          =   255
         Left            =   6360
         TabIndex        =   11
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox CharLGPDataDirCheck 
         Height          =   255
         Left            =   6360
         TabIndex        =   10
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton MagicLGPDataDirCommand 
         Caption         =   "..."
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox MagicLGPDataDirText 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   1200
         Width           =   1600
      End
      Begin VB.TextBox BattleLGPDataDirText 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   840
         Width           =   1600
      End
      Begin VB.CommandButton BattleLGPDataDirCommand 
         Caption         =   "..."
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton CharLGPDataDirCommand 
         Caption         =   "..."
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox CharLGPDataDirText 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   1600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Num frames"
         Height          =   375
         Left            =   5400
         TabIndex        =   31
         Top             =   0
         Width           =   615
      End
      Begin VB.Label InterpLabel 
         Caption         =   "Interp."
         Height          =   255
         Left            =   6240
         TabIndex        =   26
         Top             =   120
         Width           =   495
      End
      Begin VB.Label DestinationPathLabel 
         Caption         =   "Destination path"
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label SourcePathLabel 
         Caption         =   "Source path"
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   120
         Width           =   975
      End
      Begin VB.Label MagicLGPDataDirLabel 
         Caption         =   "Magic LGP"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label BattleLGPDataDirLabel 
         Caption         =   "Battle LGP"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label CharLGPDataDirLabel 
         Caption         =   "Char LGP"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "InterpolateAllAnimsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OperationCancelled As Boolean
Private Const UNIQUE_CHAR_ANIMS_COUNT = 5312
Private Const UNIQUE_BATTLE_ANIMS_COUNT = 391
Private Const UNIQUE_MAGIC_ANIMS_COUNT = 79

Private Sub BattleLGPDataDirDestCommand_Click()
    BATTLE_LGP_PATH_DEST = DirectoryToTextBox(BATTLE_LGP_PATH_DEST)
    BattleLGPDataDirDestText.Text = BATTLE_LGP_PATH_DEST
End Sub

Private Sub CancelCommand_Click()
    OperationCancelled = True
    Me.Hide
End Sub
Private Sub CharLGPDataDirCommand_Click()
    CHAR_LGP_PATH = DirectoryToTextBox(CHAR_LGP_PATH)
    CharLGPDataDirText.Text = CHAR_LGP_PATH
End Sub
Private Sub BattleLGPDataDirCommand_Click()
    BATTLE_LGP_PATH = DirectoryToTextBox(BATTLE_LGP_PATH)
    BattleLGPDataDirText.Text = BATTLE_LGP_PATH
End Sub


Private Sub CharLGPDataDirDestCommand_Click()
    CHAR_LGP_PATH_DEST = DirectoryToTextBox(CHAR_LGP_PATH_DEST)
    CharLGPDataDirDestText.Text = CHAR_LGP_PATH_DEST
End Sub

Private Sub MagicLGPDataDirCommand_Click()
    MAGIC_LGP_PATH = DirectoryToTextBox(MAGIC_LGP_PATH)
    MagicLGPDataDirText.Text = MAGIC_LGP_PATH
End Sub

Private Sub MagicLGPDataDirDestCommand_Click()
    MAGIC_LGP_PATH_DEST = DirectoryToTextBox(MAGIC_LGP_PATH_DEST)
    MagicLGPDataDirDestText.Text = MAGIC_LGP_PATH_DEST
End Sub

Private Sub Form_Load()
    If Me.Visible Then
        ResetForm
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    OperationCancelled = True
End Sub

Private Sub GoCommand_Click()
    ProgressFrame.Top = InterpolateOptionsFrame.Top + InterpolateOptionsFrame.height / 2 - ProgressFrame.height / 2
    InterpolateOptionsFrame.Visible = False
    ProgressFrame.Visible = True
    
    InterpolateAllAnimations
End Sub
Private Sub SaveConfigCommand_Click()
    DEFAULT_BATTLE_INTERP_FRAMES = NumInterpFramesBattleUpDown.value
    DEFAULT_FIELD_INTERP_FRAMES = NumInterpFramesFieldUpDown.value

    WriteCFGFile
End Sub

Public Sub ResetForm()
    ProgressFrame.Visible = False
    InterpolateOptionsFrame.Visible = True
    
    CharLGPDataDirText.Text = CHAR_LGP_PATH
    BattleLGPDataDirText.Text = BATTLE_LGP_PATH
    MagicLGPDataDirText.Text = MAGIC_LGP_PATH
    
    CharLGPDataDirDestText.Text = CHAR_LGP_PATH_DEST
    BattleLGPDataDirDestText.Text = BATTLE_LGP_PATH_DEST
    MagicLGPDataDirDestText.Text = MAGIC_LGP_PATH_DEST
    
    NumInterpFramesFieldUpDown.value = DEFAULT_FIELD_INTERP_FRAMES
    NumInterpFramesBattleUpDown.value = DEFAULT_BATTLE_INTERP_FRAMES
    
    OperationCancelled = False
End Sub

Private Sub InterpolateAllAnimations()
    Dim used_char_anims() As String
    Dim num_used_char_anims As Integer
    Dim total_unique_num_char_anims As Integer
    Dim num_anim_groups As Single
    Dim base_percentage As Single
    Dim mi As Integer
    Dim ai As Integer
    Dim aiu As Integer
    Dim foundQ As Boolean
    Dim hrc_sk As HRCSkeleton
    Dim a_anim As AAnimation
    
    Dim PI As Integer
    Dim battle_anims_packs_names(UNIQUE_BATTLE_ANIMS_COUNT - 1) As String
    Dim num_used_battle_anims_pack As Integer
    Dim aa_sk As AASkeleton
    Dim da_anims_pack As DAAnimationsPack
    Dim battle_skeleton_filename As String
    Dim limit_owner_skeleton_filename As String
    Dim anims_pack_filename As String
    
    Dim magic_anims_packs_names(UNIQUE_BATTLE_ANIMS_COUNT - 1) As String
    
    InitializeProgressBar
    
    num_anim_groups = IIf(CharLGPDataDirCheck.value = vbChecked, 1, 0)
    num_anim_groups = num_anim_groups + IIf(BattleLGPDataDirCheck.value = vbChecked, 1, 0)
    num_anim_groups = num_anim_groups + IIf(MagicLGPDataDirCheck.value = vbChecked, 1, 0)
        
    base_percentage = 0
    If CharLGPDataDirCheck.value = vbChecked Then
        num_used_char_anims = 0
        For mi = 0 To NumCharLGPRegisters - 1
            If OperationCancelled Then
                Exit For
            End If
            With CharLGPRegisters(mi)
                ReadHRCSkeleton hrc_sk, CHAR_LGP_PATH + "\" + .filename + ".HRC", False
                For ai = 0 To .NumAnims - 1
                    aiu = 0
                    While aiu < num_used_char_anims And Not foundQ
                        foundQ = (.Animations(ai) = used_char_anims(aiu))
                        aiu = aiu + 1
                    Wend
                    If Not foundQ Then
                        UpdateProgressBar (num_used_char_anims / UNIQUE_CHAR_ANIMS_COUNT) / num_anim_groups, .Animations(ai) + ".A"
                        DoEvents
                        Refresh
                        
                        ReadAAnimation a_anim, CHAR_LGP_PATH + "\" + .Animations(ai) + ".A"
                        FixAAnimation hrc_sk, a_anim
                        If a_anim.NumBones = hrc_sk.NumBones Then
                            InterpolateAAnimation hrc_sk, a_anim, NumInterpFramesFieldUpDown.value, False
                            WriteAAnimation a_anim, CHAR_LGP_PATH_DEST + "\" + .Animations(ai) + ".A"
                            num_used_char_anims = num_used_char_anims + 1
                            ReDim used_char_anims(num_used_char_anims - 1)
                            used_char_anims(num_used_char_anims - 1) = .Animations(ai)
                        End If
                    End If
                Next ai
            End With
        Next mi
        
        base_percentage = 1# / num_anim_groups
    End If
    
    If BattleLGPDataDirCheck.value = vbChecked And Not OperationCancelled Then
        anims_pack_filename = Dir(BATTLE_LGP_PATH + "\*da")

        PI = 0
        Do While anims_pack_filename > ""
            battle_anims_packs_names(PI) = anims_pack_filename
            anims_pack_filename = Dir()
            PI = PI + 1
        Loop
        
        For PI = 0 To UNIQUE_BATTLE_ANIMS_COUNT - 1
            If OperationCancelled Then
                Exit For
            End If
            
            UpdateProgressBar base_percentage + (PI / UNIQUE_BATTLE_ANIMS_COUNT) / num_anim_groups, battle_anims_packs_names(PI)
            DoEvents
            Refresh
            
            anims_pack_filename = BATTLE_LGP_PATH + "\" + battle_anims_packs_names(PI)
            battle_skeleton_filename = Left$(anims_pack_filename, Len(anims_pack_filename) - 2) + "aa"
            ReadAASkeleton battle_skeleton_filename, aa_sk, False, False
            ReadDAAnimationsPack anims_pack_filename, aa_sk.NumBones, aa_sk.NumBodyAnims, aa_sk.NumWeaponAnims, da_anims_pack
            
            InterpolateDAAnimationsPack aa_sk, da_anims_pack, NumInterpFramesBattleUpDown.value, False
            
            WriteDAAnimationsPack BATTLE_LGP_PATH_DEST + "\" + battle_anims_packs_names(PI), da_anims_pack
        Next
        
        base_percentage = base_percentage + 1# / num_anim_groups
    End If
    
    If MagicLGPDataDirCheck.value = vbChecked And Not OperationCancelled Then
        anims_pack_filename = Dir(MAGIC_LGP_PATH + "\*.a00")

        PI = 0
        Do While anims_pack_filename > ""
            magic_anims_packs_names(PI) = anims_pack_filename
            anims_pack_filename = Dir()
            PI = PI + 1
        Loop
        
        For PI = 0 To UNIQUE_MAGIC_ANIMS_COUNT - 1
            If OperationCancelled Then
                Exit For
            End If
            
            UpdateProgressBar base_percentage + (PI / UNIQUE_MAGIC_ANIMS_COUNT) / num_anim_groups, magic_anims_packs_names(PI)
            DoEvents
            Refresh
            
            anims_pack_filename = MAGIC_LGP_PATH + "\" + magic_anims_packs_names(PI)
            battle_skeleton_filename = Left$(anims_pack_filename, Len(anims_pack_filename) - 3) + "d"
            limit_owner_skeleton_filename = GetLimitCharacterFileName(magic_anims_packs_names(PI))
            If limit_owner_skeleton_filename <> "" Then
                ReadAASkeleton BATTLE_LGP_PATH + "\" + limit_owner_skeleton_filename, aa_sk, True, False
                ReadDAAnimationsPack anims_pack_filename, aa_sk.NumBones, 8, 8, da_anims_pack
            Else
                ReadMagicSkeleton battle_skeleton_filename, aa_sk, False
                ReadDAAnimationsPack anims_pack_filename, aa_sk.NumBones, aa_sk.NumBodyAnims, aa_sk.NumWeaponAnims, da_anims_pack
            End If
            
            InterpolateDAAnimationsPack aa_sk, da_anims_pack, NumInterpFramesBattleUpDown.value, False
            
            WriteDAAnimationsPack MAGIC_LGP_PATH_DEST + "\" + magic_anims_packs_names(PI), da_anims_pack
        Next
    End If
    
    If OperationCancelled Then
        MsgBox "Operation cancelled.", vbOKOnly, "Cancelled"
    Else
        UpdateProgressBar 1#, "Finished!"
        DoEvents
        Refresh
        MsgBox "Operation completed.", vbOKOnly, "Finished"
    End If
    Hide
End Sub
Private Sub InitializeProgressBar()
    Dim NewBrush As LOGBRUSH
    Dim hNewBrush As Long
    Dim hBrush As Long
    Dim oldb As Long
    NewBrush.lbColor = RGB(0, 0, 255)
    NewBrush.lbStyle = 0
    hNewBrush = CreateBrushIndirect(NewBrush)
    oldb = SelectObject(ProgressBarPicture.hdc, hNewBrush)
    DeleteObject oldb
End Sub
Private Sub UpdateProgressBar(ByVal percent As Single, ByVal filename As String)
    ProgressLabel.Caption = "Progress " + Str$(Fix(100# * percent)) + "% (" + filename + ")"
    Rectangle ProgressBarPicture.hdc, 0, 0, ProgressBarPicture.ScaleWidth * percent, ProgressBarPicture.ScaleHeight
End Sub

Private Function DirectoryToTextBox(ByVal init_val As String) As String
    ' Browse for Drive/Folder
    Dim oBrowseFolder As New cBrowseFolder
    Dim sFolder As String
    
    ' Trim and addbackslash to the current folder
    sFolder = AddBackslash(Trim$(init_val))
    'FieldDataDirText.Text = sFolder          'return to text
    'Refresh
    
    With oBrowseFolder                  'define object
        .lhWnd = Me.hWnd                'owner
        .sTitle = "Select a Drive"      'title
        .sFolder = sFolder              'initial folder
        .lFlags = BIF_RETURNONLYFSDIRS  'default
        sFolder = .ShowBrowse()         'go get it
        If Not .bCancelled Then
            DirectoryToTextBox = RTrim$(sFolder)
        Else
            DirectoryToTextBox = ""
        End If
    End With
    
    Set oBrowseFolder = Nothing
End Function

Private Function GetTotalNumberUniqueAnimations() As Integer
    Dim used_char_anims() As String
    
    Dim mi As Integer
    Dim ai As Integer
    Dim aiu As Integer
    
    Dim foundQ As Boolean
    
    GetTotalNumberUniqueAnimations = 0
    For mi = 0 To NumCharLGPRegisters - 1
        With CharLGPRegisters(mi)
            For ai = 0 To .NumAnims - 1
                aiu = 0
                While aiu < GetTotalNumberUniqueAnimations And Not foundQ
                    foundQ = (.Animations(ai) = used_char_anims(aiu))
                    aiu = aiu + 1
                Wend
                If Not foundQ Then
                    GetTotalNumberUniqueAnimations = GetTotalNumberUniqueAnimations + 1
                    ReDim used_char_anims(GetTotalNumberUniqueAnimations - 1)
                    used_char_anims(GetTotalNumberUniqueAnimations - 1) = .Animations(ai)
                End If
            Next ai
        End With
    Next mi
End Function

