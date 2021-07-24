VERSION 5.00
Begin VB.Form FF7FieldDBForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FF7 Field Data Base"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton SaveFieldDataDirCommand
      Caption         =   "Save Field Data Directory"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   4455
   End
   Begin VB.CommandButton SelectFiedlDataDirCommand
      Caption         =   "..."
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox FieldDataDirText
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Frame ModelNamesFrame
      Caption         =   "Model names"
      Height          =   2175
      Left            =   2280
      TabIndex        =   5
      ToolTipText     =   "Names used to referenced this model in field files."
      Top             =   120
      Width           =   2295
      Begin VB.Label ModelNamesLabel
         Height          =   1695
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton LoadCommand
      Caption         =   "Load animation and model"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.ComboBox ModelCombo
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "ModelCombo"
      ToolTipText     =   "Model"
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox AnimationList
      Height          =   2010
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "List of animations applied to the selected model on field files."
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label FieldDataDirLabel
      AutoSize        =   -1  'True
      Caption         =   "Field data dir"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   900
   End
   Begin VB.Line Line1
      X1              =   240
      X2              =   4560
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label AnimationLabel
      AutoSize        =   -1  'True
      Caption         =   "Animation"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   690
   End
   Begin VB.Label ModelLabel
      AutoSize        =   -1  'True
      Caption         =   "Model"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "FF7FieldDBForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub FillControls()
    Dim mi As Integer

    ModelCombo.Clear
    For mi = 0 To NumCharLGPRegisters - 1
        ModelCombo.AddItem CharLGPRegisters(mi).fileName
    Next mi

    ModelCombo.ListIndex = 0
    ModelCombo_Click

    FieldDataDirText = CHAR_LGP_PATH
End Sub

Private Sub FieldDataDirText_Change()
    CHAR_LGP_PATH = FieldDataDirText.Text
End Sub

Private Sub LoadCommand_Click()
    Dim anim_name As String
    anim_name = AnimationList.List(AnimationList.ListIndex)
    ModelEditor.OpenFF7File FieldDataDirText.Text + "\" + ModelCombo.List(ModelCombo.ListIndex) + ".HRC"
    ModelEditor.SetFieldModelAnimation FieldDataDirText.Text + "\" + anim_name + ".A"
End Sub

Private Sub ModelCombo_Click()
    Dim ai As Integer
    Dim ni As Integer
    Dim names_str As String

    With CharLGPRegisters(ModelCombo.ListIndex)
        AnimationList.Clear
        For ai = 0 To .NumAnims - 1
            AnimationList.AddItem .Animations(ai)
        Next ai

        names_str = ""
        For ni = 0 To .NumNames - 1
            names_str = names_str + .Names(ni) + ", "
        Next ni
        ModelNamesLabel.Caption = names_str
    End With
End Sub

Private Sub SaveFieldDataDirCommand_Click()
    WriteCFGFile
End Sub

Private Sub SelectFiedlDataDirCommand_Click()
    ' Browse for Drive/Folder
    Dim oBrowseFolder As New cBrowseFolder
    Dim sFolder As String

    ' Trim and addbackslash to the current folder
    sFolder = AddBackslash(Trim$(FieldDataDirText.Text))
    FieldDataDirText.Text = sFolder          'return to text
    Refresh

    With oBrowseFolder                  'define object
        .lhWnd = Me.hWnd                'owner
        .sTitle = "Select a Drive"      'title
        .sFolder = sFolder              'initial folder
        .lFlags = BIF_RETURNONLYFSDIRS  'default
        sFolder = .ShowBrowse()         'go get it
        If Not .bCancelled Then
            FieldDataDirText.Text = RTrim$(sFolder)
        End If
    End With

    Set oBrowseFolder = Nothing
End Sub
