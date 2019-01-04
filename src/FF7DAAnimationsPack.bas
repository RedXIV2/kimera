Attribute VB_Name = "FF7DAAnimationsPack"
Option Explicit
Type DAAnimationsPack
    'Battle animations notes by L.Spiro and Qhimm:
    'http://wiki.qhimm.com/FF7/Battle/Battle_Animation_(PC)
    'Weapon animations notes by seb:
    'http://forums.qhimm.com/index.php?topic=7185.0
    NumAnimations As Long
    NumBodyAnimations As Long
    NumWeaponAnimations As Long
    BodyAnimations() As DAAnimation
    WeaponAnimations() As DAAnimation
End Type
Sub ReadDAAnimationsPack(ByVal filename As String, ByVal NumBones As Integer, ByVal NumBodyAnimations As Long, ByVal NumWeaponAnimations As Long, ByRef AnimationsPack As DAAnimationsPack)
    Dim fileNumber As Integer
    Dim ai As Integer
    Dim offset As Long
    
    On Error GoTo ErrHandRead
    'Debug.Print "Loadng animations pack " + fileName
    
    'Debug.Print "Reading pack "; fileName
    
    fileNumber = FreeFile
    Open filename For Binary As fileNumber
    
    With AnimationsPack
        Get fileNumber, 1, .NumAnimations
        '.NumAnimations = 1
        'NumBodyAnimations = 1
        .NumBodyAnimations = NumBodyAnimations
        'NumWeaponAnimations = 1
        .NumWeaponAnimations = NumWeaponAnimations
        
        ReDim .BodyAnimations(NumBodyAnimations)
        ReDim .WeaponAnimations(NumWeaponAnimations)
        offset = 5
        ''Debug.Print "Loading "; .NumAnimations; " animations."
        For ai = 0 To NumBodyAnimations - 1
            ''Debug.Print "anim "; ai
            'Debug.Print "Body Animation "; Str$(ai)
            ReadDAAnimation fileNumber, offset, IIf(NumBones > 1, NumBones + 1, 1), .BodyAnimations(ai)
            'offset = offset + .Animations(ai).BlockLength
        Next ai
        
        For ai = 0 To NumWeaponAnimations - 1
            ''Debug.Print "anim "; ai
            'Debug.Print "Weapon Animation "; Str$(ai)
            ReadDAAnimation fileNumber, offset, 1, .WeaponAnimations(ai)
            'offset = offset + .Animations(ai).BlockLength
        Next ai
    End With
    
    Close fileNumber
    Exit Sub
ErrHandRead:
    If Err < 1000 Then
        'Debug.Print "Error reading DA file!!!"
        MsgBox "Error reading DA file " + filename + "!!!", vbOKOnly, "Error reading DA file"
        CreateEmptyDAAnimationsPack AnimationsPack, NumBones
    Else
        'Debug.Print "Error reading animation "; ai
        MsgBox "Error reading animation " + Str$(ai) + "(frame " + Str$(Err - 1000) + ") from DA file " + filename + "!!!", vbOKOnly, "Error reading DA file"
        AnimationsPack.NumAnimations = ai - 1
        'Resume Next
    End If
End Sub
Sub WriteDAAnimationsPack(ByVal filename As String, ByRef AnimationsPack As DAAnimationsPack)
    Dim fileNumber As Integer
    Dim ai As Integer
    Dim offset As Long
    
    On Error GoTo ErrHandRead
    
    'Since we're using signed data there is no way we can store values outside the [-180º, 180º] (shoudln't matter though, due to angular equivalences)
    'Normalize just to be safe
    NormalizeDAAnimationsPack AnimationsPack
    
    'Debug.Print "Writting pack "; fileName
    
    fileNumber = FreeFile
    Open filename For Output As fileNumber
    Close fileNumber
    Open filename For Binary As fileNumber
    
    With AnimationsPack
        Put fileNumber, 1, .NumAnimations
        
        offset = 5
        For ai = 0 To .NumBodyAnimations - 1
            WriteDAAnimation fileNumber, offset, .BodyAnimations(ai)
            'offset = offset + .BodyAnimations(ai).BlockLength
        Next ai
        For ai = 0 To .NumWeaponAnimations - 1
            'Debug.Print "Weapon Animation "; Str$(ai)
            If ai = 7 Then
                ai = ai
            End If
            WriteDAAnimation fileNumber, offset, .WeaponAnimations(ai)
            'offset = offset + .WeaponAnimations(ai).BlockLength
        Next ai
    End With
    
    Close fileNumber
    Exit Sub
ErrHandRead:
    'Debug.Print "Error writting DA file!!!"
    MsgBox "Error writting DA file " + filename + "!!!", vbOKOnly, "Error writting"
End Sub
Sub CreateEmptyDAAnimationsPack(ByRef AnimationsPack As DAAnimationsPack, ByVal NumBones As Integer)
    With AnimationsPack
        .NumAnimations = 1
        .NumBodyAnimations = 1
        .NumWeaponAnimations = 0
        ReDim .BodyAnimations(1)
        ReDim .WeaponAnimations(0)
        CreateEmptyDAAnimationsPackAnimation .BodyAnimations(0), NumBones
    End With
End Sub

Sub NormalizeDAAnimationsPack(ByRef AnimationsPack As DAAnimationsPack)
    Dim ai As Integer
    Dim BI As Integer
    Dim num_bones As Integer
    
    For ai = 0 To AnimationsPack.NumBodyAnimations - 1
        If AnimationsPack.BodyAnimations(ai).BlockLength >= 11 And AnimationsPack.BodyAnimations(ai).NumFrames2 > 0 Then
            num_bones = UBound(AnimationsPack.BodyAnimations(ai).Frames(0).Bones) + 1
            For BI = 0 To num_bones - 1
                With AnimationsPack.BodyAnimations(ai).Frames(0).Bones(BI)
                    .alpha = NormalizeAngle180(.alpha)
                    .Beta = NormalizeAngle180(.Beta)
                    .Gamma = NormalizeAngle180(.Gamma)
                End With
            Next BI
            NormalizeDAAnimationsPackAnimation AnimationsPack.BodyAnimations(ai)
        End If
    Next ai
    For ai = 0 To AnimationsPack.NumWeaponAnimations - 1
        If AnimationsPack.WeaponAnimations(ai).BlockLength >= 11 And AnimationsPack.WeaponAnimations(ai).NumFrames2 > 0 Then
            With AnimationsPack.WeaponAnimations(ai).Frames(0).Bones(0)
                .alpha = NormalizeAngle180(.alpha)
                .Beta = NormalizeAngle180(.Beta)
                .Gamma = NormalizeAngle180(.Gamma)
            End With
            NormalizeDAAnimationsPackAnimation AnimationsPack.WeaponAnimations(ai)
        End If
    Next ai
End Sub

Sub CheckWriteDAAnimationsPack(ByVal filename As String, ByRef AnimationsPack As DAAnimationsPack)
    Dim fileNumber As Integer
    Dim ai As Integer
    Dim offset As Long
    
    Dim num_animations As Long
    Dim num_body_animations As Long
    Dim num_weapon_animations As Long
    
    Dim body_animations() As DAAnimation
    Dim weapon_animations() As DAAnimation
    
    fileNumber = FreeFile
    Open filename For Binary As fileNumber
    
    With AnimationsPack
        Get fileNumber, 1, num_animations
        If num_animations <> .NumAnimations Then
            Debug.Print "Error"
        End If
        
        ReDim body_animations(.NumBodyAnimations)
        ReDim weapon_animations(.NumWeaponAnimations)
        offset = 5

        For ai = 0 To .NumBodyAnimations - 1
            CheckWriteDAAnimation fileNumber, offset, .BodyAnimations(ai)
        Next ai
        
        For ai = 0 To .NumWeaponAnimations - 1
            CheckWriteDAAnimation fileNumber, offset, .WeaponAnimations(ai)
        Next ai
    End With
    
    Close fileNumber
End Sub
