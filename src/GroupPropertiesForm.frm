VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form GroupPropertiesForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Group properties"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame RenderStateFrame 
      Caption         =   "Render state"
      Height          =   3855
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   3735
      Begin VB.CheckBox ShadeModeLightedCheck 
         Caption         =   "Lighted"
         Height          =   255
         Left            =   2160
         TabIndex        =   30
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CheckBox AlphaBlendTrueCheck 
         Caption         =   "True"
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   3120
         Width           =   855
      End
      Begin VB.CheckBox DepthMaskTrueCheck 
         Caption         =   "True"
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox DepthTestTrueCheck 
         Caption         =   "True"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox CullFaceClockWiseCheck 
         Caption         =   "Cull back-facing"
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox NoCullTrueCheck 
         Caption         =   "True"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox LinearFilterTrueCheck 
         Caption         =   "True"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   1320
         Width           =   975
      End
      Begin VB.CheckBox TexturedTrueCheck 
         Caption         =   "True"
         Height          =   195
         Left            =   2160
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox WireframeTrueCheck 
         Caption         =   "True"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox V_SHADEMODE_Check 
         Caption         =   "V_SHADEMODE"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CheckBox V_ALPHABLEND_Check 
         Caption         =   "V_ALPHABLEND"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   2175
      End
      Begin VB.CheckBox V_DEPTHMASK_Check 
         Caption         =   "V_DEPTHMASK"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CheckBox V_DEPTHTEST_Check 
         Caption         =   "V_DEPTHTEST"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CheckBox V_CULLFACE_Check 
         Caption         =   "V_CULLFACE"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CheckBox V_NOCULL_Check 
         Caption         =   " V_NOCULL"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox V_LINEARFILTER_Check 
         Caption         =   "V_LINEARFILTER"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CheckBox V_TEXTURE_Check 
         Caption         =   "V_TEXTURE"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox V_WIREFRAME_Check 
         Caption         =   "V_WIREFRAME"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label RSValueLabel 
         Caption         =   "RS Value:"
         Height          =   255
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label EnableChangeLabel 
         Caption         =   "Enable RS change:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComCtl2.UpDown TextureIdUpDown 
      Height          =   285
      Left            =   3600
      TabIndex        =   8
      Top             =   6000
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "TextureIdText"
      BuddyDispid     =   196630
      OrigLeft        =   1680
      OrigTop         =   2160
      OrigRight       =   1920
      OrigBottom      =   2415
      Max             =   9999
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox TextureIdText 
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton ApplyButton 
      Caption         =   "Apply"
      Height          =   1815
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame GroupOpacityFrame 
      Caption         =   "Blending mode"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.OptionButton BlendingNoneOption 
         Caption         =   "None"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "No blending"
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton BlendUnkownOption 
         Caption         =   "Unknown (broken?)"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "¿?"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton BlendSubstractiveOption 
         Caption         =   "Substractive"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "should be destination color - source color"
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton BlendAdditiveOption 
         Caption         =   "Additive"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "source color + destination color"
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton BlendAverageOption 
         Caption         =   "Average"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "source color / 2 + destination color / 2"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label TextureIdLabel 
      Caption         =   "Texture Id"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   735
   End
End
Attribute VB_Name = "GroupPropertiesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedGroup As Integer
Public Sub SetSelectedGroup(ByVal Group_index As Integer)
    SelectedGroup = Group_index
    Me.Caption = "Editing group " + Str$(SelectedGroup)
    
    TextureIdUpDown.Enabled = _
        (EditedPModel.Groups(SelectedGroup).texFlag = 1)
    
    TextureIdText.Enabled = TextureIdUpDown.Enabled
    
    TextureIdText.Text = EditedPModel.Groups(SelectedGroup).TexID
    
    Select Case EditedPModel.hundrets(SelectedGroup).blend_mode
        Case 0:
            BlendAverageOption.value = True
        Case 1:
            BlendAdditiveOption.value = True
        Case 2:
            BlendSubstractiveOption.value = True
        Case 3:
            BlendUnkownOption.value = True
        Case 4:
            BlendingNoneOption.value = True
    End Select
    
    Dim change_render_state_values As Long
    Dim render_state_values As Long
    
    change_render_state_values = EditedPModel.hundrets(SelectedGroup).field_8
    render_state_values = EditedPModel.hundrets(SelectedGroup).field_C
    
    V_WIREFRAME_Check.value = IIf((change_render_state_values And &H1&) = 0, vbUnchecked, vbChecked)
    V_TEXTURE_Check.value = IIf((change_render_state_values And &H2&) = 0, vbUnchecked, vbChecked)
    V_LINEARFILTER_Check.value = IIf((change_render_state_values And &H4&) = 0, vbUnchecked, vbChecked)
    V_NOCULL_Check.value = IIf((change_render_state_values And &H4000&) = 0, vbUnchecked, vbChecked)
    V_CULLFACE_Check.value = IIf((change_render_state_values And &H2000&) = 0, vbUnchecked, vbChecked)
    V_DEPTHTEST_Check.value = IIf((change_render_state_values And &H8000&) = 0, vbUnchecked, vbChecked)
    V_DEPTHMASK_Check.value = IIf((change_render_state_values And &H10000) = 0, vbUnchecked, vbChecked)
    V_ALPHABLEND_Check.value = IIf((change_render_state_values And &H400&) = 0, vbUnchecked, vbChecked)
    V_SHADEMODE_Check.value = IIf((change_render_state_values And &H20000) = 0, vbUnchecked, vbChecked)
    
    WireframeTrueCheck.value = IIf((render_state_values And &H1&) = 0, vbUnchecked, vbChecked)
    TexturedTrueCheck.value = IIf((render_state_values And &H2&) = 0, vbUnchecked, vbChecked)
    LinearFilterTrueCheck.value = IIf((render_state_values And &H4&) = 0, vbUnchecked, vbChecked)
    NoCullTrueCheck.value = IIf((render_state_values And &H4000&) = 0, vbUnchecked, vbChecked)
    CullFaceClockWiseCheck.value = IIf((render_state_values And &H2000&) = 0, vbUnchecked, vbChecked)
    DepthTestTrueCheck.value = IIf((render_state_values And &H8000&) = 0, vbUnchecked, vbChecked)
    DepthMaskTrueCheck.value = IIf((render_state_values And &H10000) = 0, vbUnchecked, vbChecked)
    AlphaBlendTrueCheck.value = IIf((render_state_values And &H400&) = 0, vbUnchecked, vbChecked)
    ShadeModeLightedCheck.value = IIf((render_state_values And &H20000) = 0, vbUnchecked, vbChecked)
    
    SetupRSValuesEnabled
End Sub
Private Sub ApplyButton_Click()
    With EditedPModel.hundrets(SelectedGroup)
        If BlendAverageOption.value = True Then
            .blend_mode = 0
        ElseIf BlendAdditiveOption.value = True Then
            .blend_mode = 1
        ElseIf BlendSubstractiveOption.value = True Then
            .blend_mode = 2
        ElseIf BlendUnkownOption.value = True Then
            .blend_mode = 3
        ElseIf BlendingNoneOption.value = True Then
            .blend_mode = 4
        End If
        
        If EditedPModel.Groups(SelectedGroup).texFlag = 1 Then
            EditedPModel.Groups(SelectedGroup).TexID = TextureIdUpDown.value
            .TexID = TextureIdUpDown.value
        End If
            
        Dim change_render_state_values As Long
        Dim render_state_values As Long
        
        change_render_state_values = .field_8
        render_state_values = .field_C
            
        Dim mask As Long
        Dim inv_mask_full As Long
        
        inv_mask_full = Not (&H1& Or &H2& Or &H4& Or &H4000& Or &H2000& Or &H8000& Or &H10000 Or &H400& Or &H20000)
        
        mask = IIf(V_WIREFRAME_Check.value = vbChecked, &H1&, &H0&) Or _
                IIf(V_TEXTURE_Check.value = vbChecked, &H2&, &H0&) Or _
                IIf(V_LINEARFILTER_Check.value = vbChecked, &H4&, &H0&) Or _
                IIf(V_NOCULL_Check.value = vbChecked, &H4000&, &H0&) Or _
                IIf(V_CULLFACE_Check.value = vbChecked, &H2000&, &H0&) Or _
                IIf(V_DEPTHTEST_Check.value = vbChecked, &H8000&, &H0&) Or _
                IIf(V_DEPTHMASK_Check.value = vbChecked, &H10000, &H0&) Or _
                IIf(V_ALPHABLEND_Check.value = vbChecked, &H400&, &H0&) Or _
                IIf(V_SHADEMODE_Check.value = vbChecked, &H20000, &H0&)
        change_render_state_values = (change_render_state_values And inv_mask_full) Or mask

                                        
        mask = IIf(WireframeTrueCheck.value = vbChecked, &H1&, &H0&) Or _
                IIf(TexturedTrueCheck.value = vbChecked, &H2&, &H0&) Or _
                IIf(LinearFilterTrueCheck.value = vbChecked, &H4&, &H0&) Or _
                IIf(NoCullTrueCheck.value = vbChecked, &H4000&, &H0&) Or _
                IIf(CullFaceClockWiseCheck.value = vbChecked, &H2000&, &H0&) Or _
                IIf(DepthTestTrueCheck.value = vbChecked, &H8000&, &H0&) Or _
                IIf(DepthMaskTrueCheck.value = vbChecked, &H10000, &H0&) Or _
                IIf(AlphaBlendTrueCheck.value = vbChecked, &H400&, &H0&) Or _
                IIf(ShadeModeLightedCheck.value = vbChecked, &H20000, &H0&)
        render_state_values = (render_state_values And inv_mask_full) Or mask
                                        
        .field_8 = change_render_state_values
        .field_C = render_state_values
        
        V_WIREFRAME_Check.value = IIf((change_render_state_values And &H1&) = 0, vbUnchecked, vbChecked)
        V_TEXTURE_Check.value = IIf((change_render_state_values And &H2&) = 0, vbUnchecked, vbChecked)
        V_LINEARFILTER_Check.value = IIf((change_render_state_values And &H4&) = 0, vbUnchecked, vbChecked)
        V_NOCULL_Check.value = IIf((change_render_state_values And &H4000&) = 0, vbUnchecked, vbChecked)
        V_CULLFACE_Check.value = IIf((change_render_state_values And &H2000&) = 0, vbUnchecked, vbChecked)
        V_DEPTHTEST_Check.value = IIf((change_render_state_values And &H8000&) = 0, vbUnchecked, vbChecked)
        V_DEPTHMASK_Check.value = IIf((change_render_state_values And &H10000) = 0, vbUnchecked, vbChecked)
        V_ALPHABLEND_Check.value = IIf((change_render_state_values And &H400&) = 0, vbUnchecked, vbChecked)
        V_SHADEMODE_Check.value = IIf((change_render_state_values And &H20000) = 0, vbUnchecked, vbChecked)
        
        SetupRSValuesEnabled
        
        WireframeTrueCheck.value = IIf((render_state_values And &H1&) = 0, vbUnchecked, vbChecked)
        TexturedTrueCheck.value = IIf((render_state_values And &H2&) = 0, vbUnchecked, vbChecked)
        LinearFilterTrueCheck.value = IIf((render_state_values And &H4&) = 0, vbUnchecked, vbChecked)
        NoCullTrueCheck.value = IIf((render_state_values And &H4000&) = 0, vbUnchecked, vbChecked)
        CullFaceClockWiseCheck.value = IIf((render_state_values And &H2000&) = 0, vbUnchecked, vbChecked)
        DepthTestTrueCheck.value = IIf((render_state_values And &H8000&) = 0, vbUnchecked, vbChecked)
        DepthMaskTrueCheck.value = IIf((render_state_values And &H10000) = 0, vbUnchecked, vbChecked)
        AlphaBlendTrueCheck.value = IIf((render_state_values And &H400&) = 0, vbUnchecked, vbChecked)
        ShadeModeLightedCheck.value = IIf((render_state_values And &H20000) = 0, vbUnchecked, vbChecked)
    End With
End Sub



Private Sub TextureIdText_Change()
    If IsNumeric(TextureIdText.Text) Then
        If TextureIdText.Text >= TextureIdUpDown.Min And TextureIdText.Text <= TextureIdUpDown.max Then
            TextureIdUpDown.value = TextureIdText.Text
        Else
            Beep
        End If
    Else
        Beep
    End If
End Sub

Private Sub V_ALPHABLEND_Check_Click()
    SetupRSValuesEnabled
End Sub

Private Sub V_CULLFACE_Check_Click()
    SetupRSValuesEnabled
End Sub

Private Sub V_DEPTHMASK_Check_Click()
    SetupRSValuesEnabled
End Sub

Private Sub V_DEPTHTEST_Check_Click()
    SetupRSValuesEnabled
End Sub

Private Sub V_LINEARFILTER_Check_Click()
    SetupRSValuesEnabled
End Sub

Private Sub V_NOCULL_Check_Click()
    SetupRSValuesEnabled
End Sub

Private Sub V_SHADEMODE_Check_Click()
    SetupRSValuesEnabled
End Sub

Private Sub V_TEXTURE_Check_Click()
    SetupRSValuesEnabled
End Sub

Private Sub V_WIREFRAME_Check_Click()
    SetupRSValuesEnabled
End Sub

Private Sub SetupRSValuesEnabled()
    WireframeTrueCheck.Enabled = (V_WIREFRAME_Check.value = vbChecked)
    TexturedTrueCheck.Enabled = (V_TEXTURE_Check.value = vbChecked)
    LinearFilterTrueCheck.Enabled = (V_LINEARFILTER_Check.value = vbChecked)
    NoCullTrueCheck.Enabled = (V_NOCULL_Check.value = vbChecked)
    CullFaceClockWiseCheck.Enabled = (V_CULLFACE_Check.value = vbChecked)
    DepthTestTrueCheck.Enabled = (V_DEPTHTEST_Check.value = vbChecked)
    DepthMaskTrueCheck.Enabled = (V_DEPTHMASK_Check.value = vbChecked)
    AlphaBlendTrueCheck.Enabled = (V_ALPHABLEND_Check.value = vbChecked)
    ShadeModeLightedCheck.Enabled = (V_SHADEMODE_Check.value = vbChecked)
End Sub
