Attribute VB_Name = "FF7HRCSkeleton"
Option Explicit
Type HRCSkeleton
    filename As String
    name As String
    NumBones As Integer
    Bones() As HRCBone
End Type
Sub ReadHRCSkeleton(ByRef obj As HRCSkeleton, ByVal filename As String, ByVal load_geometryQ As Boolean)
    Dim BI As Integer
    Dim n_bones As Integer
    Dim fileNumber As Integer
    Dim textures_pool() As TEXTexture

    Dim line As String

    On Error GoTo errorH

    ChDir GetPathFromString(filename)

    fileNumber = FreeFile
    Open filename For Input As fileNumber
    obj.filename = TrimPath(filename)

    'Skip all empty lines or commments

    'Allways ":HEADER_BLOCK 2".
    Do
        Line Input #fileNumber, line
    Loop Until Left$(line, 1) = ":"

    'Skeleton name.
    Do
        Line Input #fileNumber, line
    Loop Until Left$(line, 1) = ":"
    obj.name = Mid$(line, 10, Len(line))

    'Number of bones.
    Do
        Line Input #fileNumber, line
    Loop Until Left$(line, 1) = ":"
    n_bones = Mid$(line, 7, Len(line))

    'Objects without a skeleton
    If n_bones = 0 Then n_bones = 1
    obj.NumBones = n_bones

    ReDim obj.Bones(n_bones * 2)

    For BI = 0 To obj.NumBones - 1
        ReadHRCBone fileNumber, obj.Bones(BI), textures_pool, load_geometryQ
    Next BI

    Close fileNumber
    Exit Sub
errorH:
    'Debug.Print "Error reading HRC file!!!"
    MsgBox "Error reading HRC file " + filename + "!!!", vbOKOnly, "Error reading"
End Sub
Sub WriteHRCSkeleton(ByRef obj As HRCSkeleton, ByVal filename As String)
    Dim BI As Integer
    Dim pi As Integer
    Dim fileNumber As Integer

    On Error GoTo errorH

    ChDir GetPathFromString(filename)

    fileNumber = FreeFile
    Open filename For Output As fileNumber

    With obj
        .filename = TrimPath(filename)
        Print #fileNumber, ":HEADER_BLOCK 2"
        Print #fileNumber, ":SKELETON" + .name
        Print #fileNumber, ":BONES " + Right$(Str$(.NumBones), Len(Str$(.NumBones)) - 1)

        For BI = 0 To .NumBones - 1
            WriteHRCBone fileNumber, .Bones(BI)
        Next BI
    End With

    Close fileNumber
    Exit Sub
errorH:
    'Debug.Print "Error writting HRC file!!!"
    MsgBox "Error writing HRC file " + filename + "!!!", vbOKOnly, "Error writing"
End Sub
Sub CreateDListsFromHRCSkeleton(ByRef obj As HRCSkeleton)
    Dim BI As Integer

    With obj
        For BI = 0 To .NumBones - 1
            CreateDListsFromHRCSkeletonBone .Bones(BI)
        Next BI
    End With
End Sub
Sub FreeHRCSkeletonResources(ByRef obj As HRCSkeleton)
    Dim BI As Integer

    With obj
        For BI = 0 To .NumBones - 1
            FreeHRCBoneResources .Bones(BI)
        Next BI
    End With
End Sub
Sub DrawHRCSkeleton(ByRef obj As HRCSkeleton, ByRef Frame As AFrame, ByVal UseDLists As Boolean)
    Dim BI As Integer
    Dim last_joint As String
    Dim joint_stack() As String
    Dim jsp As Integer
    Dim rot_mat(16) As Double

    glMatrixMode GL_MODELVIEW

    glPushMatrix
    With Frame
        glTranslated .RootTranslationX, 0, 0
        glTranslated 0, -.RootTranslationY, 0
        glTranslated 0, 0, .RootTranslationZ

        BuildRotationMatrixWithQuaternions .RootRotationAlpha, .RootRotationBeta, _
            .RootRotationGamma, rot_mat
        glMultMatrixd rot_mat(0)
    End With

    ReDim joint_stack(obj.NumBones)
    jsp = 0

    joint_stack(jsp) = obj.Bones(0).joint_f
    For BI = 0 To obj.NumBones - 1
        While Not (obj.Bones(BI).joint_f = joint_stack(jsp)) And jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend
        If (jsp = 0) Then
            SetDefaultOGLRenderState
        End If

        glPushMatrix

        'glRotated Frame.Rotations(bi).beta, 0#, 1#, 0#
        'glRotated Frame.Rotations(bi).alpha, 1#, 0#, 0#
        'glRotated Frame.Rotations(bi).gamma, 0#, 0#, 1#
        With Frame.Rotations(BI)
            BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
        End With
        glMultMatrixd rot_mat(0)

        DrawHRCBone obj.Bones(BI), UseDLists

        glTranslated 0, 0, -obj.Bones(BI).length

        jsp = jsp + 1
        joint_stack(jsp) = obj.Bones(BI).joint_i
    Next BI

    While jsp > 0
        glPopMatrix
        jsp = jsp - 1
    Wend
    glPopMatrix
End Sub
Sub DrawHRCSkeletonBones(ByRef obj As HRCSkeleton, ByRef Frame As AFrame)
    Dim BI As Integer
    Dim last_joint As String
    Dim joint_stack() As String
    Dim jsp As String
    Dim rot_mat(16) As Double

    glMatrixMode GL_MODELVIEW

    glPushMatrix
    With Frame
        glTranslated .RootTranslationX, 0, 0
        glTranslated 0, -.RootTranslationY, 0
        glTranslated 0, 0, .RootTranslationZ

        BuildRotationMatrixWithQuaternions .RootRotationAlpha, .RootRotationBeta, _
            .RootRotationGamma, rot_mat
        glMultMatrixd rot_mat(0)
    End With

    ReDim joint_stack(obj.NumBones)
    jsp = 0

    glPointSize 5

    joint_stack(jsp) = obj.Bones(0).joint_f
    For BI = 0 To obj.NumBones - 1
        While Not (obj.Bones(BI).joint_f = joint_stack(jsp)) And jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend
        glPushMatrix

        'glRotated Frame.Rotations(bi).Beta, 0#, 1#, 0#
        'glRotated Frame.Rotations(bi).Alpha, 1#, 0#, 0#
        'glRotated Frame.Rotations(bi).Gamma, 0#, 0#, 1#
        With Frame.Rotations(BI)
            BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
        End With
        glMultMatrixd rot_mat(0)

        glBegin GL_POINTS
            glVertex3f 0, 0, 0
            glVertex3f 0, 0, -obj.Bones(BI).length
        glEnd

        glBegin GL_LINES
            glVertex3f 0, 0, 0
            glVertex3f 0, 0, -obj.Bones(BI).length
        glEnd

        glTranslated 0, 0, -obj.Bones(BI).length

        jsp = jsp + 1
        joint_stack(jsp) = obj.Bones(BI).joint_i
    Next BI

    While jsp > 0
        glPopMatrix
        jsp = jsp - 1
    Wend
    glPopMatrix
End Sub
Sub CreateCompatibleHRCAAnimation(ByRef obj As HRCSkeleton, ByRef AAnim As AAnimation)
    Dim BI As Integer
    Dim last_joint As String
    Dim joint_stack() As String
    Dim jsp As String

    Dim StageIndex As Integer

    Dim HipArmAngle As Single
    Dim c1 As Single
    Dim c2 As Single

    ReDim joint_stack(obj.NumBones)
    jsp = 0

    joint_stack(jsp) = obj.Bones(0).joint_f

    ReDim AAnim.Frames(0)
    ReDim AAnim.Frames(0).Rotations(obj.NumBones)

    AAnim.NumFrames = 1
    AAnim.Frames(0).Rotations(0).Gamma = 180
    AAnim.Frames(0).Rotations(0).alpha = 90

    For BI = 1 To obj.NumBones - 1
        While Not (obj.Bones(BI).joint_f = joint_stack(jsp)) And jsp > 0
            jsp = jsp - 1
        Wend

        If (StageIndex < 7 And obj.Bones(BI).joint_f = "hip") Or (StageIndex >= 7 And obj.Bones(BI).joint_f = "root") Then
            StageIndex = StageIndex + 1
        End If
        ''Debug.Print jsp

        With AAnim.Frames(0).Rotations(BI)
            Select Case StageIndex
                Case 1:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                    If BI = 2 Then StageIndex = 2
                Case 2:
                    .alpha = -145
                    .Beta = 0
                    .Gamma = 0
                    StageIndex = 3
                Case 3:
                    If jsp > 1 Then
                        .alpha = 0
                        .Beta = 0
                        .Gamma = 0
                    Else
                        .alpha = 0
                        c1 = obj.Bones(1).length * 0.9
                        If c1 > obj.Bones(BI).length Then c1 = obj.Bones(BI).length - obj.Bones(BI).length * 0.01
                        c2 = Sqr(obj.Bones(BI).length ^ 2 - c1 ^ 2)
                        HipArmAngle = Atn(c2 / c1) / PI_180
                        .Beta = -HipArmAngle
                        .Gamma = 0
                        StageIndex = 5
                    End If
                Case 4:
                    .alpha = 0
                    c1 = obj.Bones(1).length * 0.9
                    If c1 > obj.Bones(BI).length Then c1 = obj.Bones(BI).length - obj.Bones(BI).length * 0.01
                    c2 = Sqr(obj.Bones(BI).length ^ 2 - c1 ^ 2)
                    HipArmAngle = Atn(c2 / c1) / PI_180
                    .Beta = -HipArmAngle
                    .Gamma = 0
                    StageIndex = 5
                Case 5:
                    .alpha = 0
                    .Beta = -90 + HipArmAngle
                    .Gamma = 180
                    StageIndex = 6
                Case 6:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                Case 7:
                    .alpha = 0
                    .Beta = HipArmAngle
                    .Gamma = 0
                    StageIndex = 8
                Case 8:
                    .alpha = 0
                    .Beta = -HipArmAngle + 90
                    .Gamma = 180
                    StageIndex = 9
                Case 9:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                Case 10:
                    .alpha = 0
                    .Beta = 90
                    .Gamma = 90
                    StageIndex = 11
                Case 11:
                    .alpha = 0
                    .Beta = 60
                    .Gamma = 0
                    StageIndex = 12
                Case 12:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                    StageIndex = 13
                Case 13:
                    .alpha = -90
                    .Beta = 0
                    .Gamma = 0
                Case 14:
                    .alpha = 0
                    .Beta = -90
                    .Gamma = -90
                    StageIndex = 15
                Case 15:
                    .alpha = 0
                    .Beta = -60
                    .Gamma = 0
                    StageIndex = 16
                Case 16:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                    StageIndex = 17
                Case 17:
                    .alpha = -90
                    .Beta = 0
                    .Gamma = 0
                Case 18:
                    .alpha = 90
                    .Beta = 0
                    .Gamma = 0
                Case Else
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
            End Select
        End With
        jsp = jsp + 1
        joint_stack(jsp) = obj.Bones(BI).joint_i
    Next BI
End Sub
Sub SetCameraHRCSkeleton(ByRef obj As HRCSkeleton, ByVal cx As Single, ByVal cy As Single, ByVal CZ As Single, ByVal alpha As Single, ByVal Beta As Single, ByVal Gamma As Single, ByVal redX As Single, ByVal redY As Single, ByVal redZ As Single)
    Dim width As Integer
    Dim height As Integer
    Dim vp(4) As Long

    glGetIntegerv GL_VIEWPORT, vp(0)
    width = vp(2)
    height = vp(3)

    glMatrixMode GL_PROJECTION
    glLoadIdentity
    gluPerspective 60, width / height, max(0.1 - CZ, 0.1), 10000 - CZ 'max(0.1 - CZ, 0.1), ComputeHRCDiameter(obj) * 2 - CZ

    Dim f_start As Single
    Dim f_end As Single
    f_start = 500 - CZ
    f_end = 10000 - CZ
    glFogfv GL_FOG_START, f_start
    glFogfv GL_FOG_END, f_end

    glMatrixMode GL_MODELVIEW
    glLoadIdentity

    glTranslatef cx, cy, CZ - ComputeHRCDiameter(obj)

    glRotatef Beta, 1#, 0#, 0#
    glRotatef alpha, 0#, 1#, 0#
    glRotatef Gamma, 0#, 0#, 1#

    glScalef redX, redY, redZ
End Sub
Sub ComputeHRCBoundingBox(ByRef obj As HRCSkeleton, ByRef Frame As AFrame, ByRef p_min_HRC As Point3D, ByRef p_max_HRC As Point3D)
    Dim last_joint As String
    Dim joint_stack() As String
    Dim jsp As Integer

    ReDim joint_stack(obj.NumBones)
    jsp = 0

    joint_stack(jsp) = obj.Bones(0).joint_f

    Dim rot_mat(16) As Double
    Dim MV_matrix(16) As Double
    Dim BI As Integer

    Dim p_max_bone As Point3D
    Dim p_min_bone As Point3D

    Dim p_max_bone_trans As Point3D
    Dim p_min_bone_trans As Point3D

    p_max_HRC.x = -INFINITY_SINGLE
    p_max_HRC.y = -INFINITY_SINGLE
    p_max_HRC.z = -INFINITY_SINGLE

    p_min_HRC.x = INFINITY_SINGLE
    p_min_HRC.y = INFINITY_SINGLE
    p_min_HRC.z = INFINITY_SINGLE

    glMatrixMode GL_MODELVIEW
    glPushMatrix
    glLoadIdentity
    With Frame
        glTranslated .RootTranslationX, 0, 0
        glTranslated 0, -.RootTranslationY, 0
        glTranslated 0, 0, .RootTranslationZ

        BuildRotationMatrixWithQuaternions .RootRotationAlpha, .RootRotationBeta, _
            .RootRotationGamma, rot_mat
        glMultMatrixd rot_mat(0)
    End With
    For BI = 0 To obj.NumBones - 1
        While Not (obj.Bones(BI).joint_f = joint_stack(jsp)) And jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend
        glPushMatrix

        BuildRotationMatrixWithQuaternions Frame.Rotations(BI).alpha, Frame.Rotations(BI).Beta, _
            Frame.Rotations(BI).Gamma, rot_mat
        glMultMatrixd rot_mat(0)

        ComputeHRCBoneBoundingBox obj.Bones(BI), p_min_bone, p_max_bone

        glGetDoublev GL_MODELVIEW_MATRIX, MV_matrix(0)

        ComputeTransformedBoxBoundingBox MV_matrix, p_min_bone, p_max_bone, _
            p_min_bone_trans, p_max_bone_trans

        With p_max_bone_trans
            If p_max_HRC.x < .x Then p_max_HRC.x = .x
            If p_max_HRC.y < .y Then p_max_HRC.y = .y
            If p_max_HRC.z < .z Then p_max_HRC.z = .z
        End With

        With p_min_bone_trans
            If p_min_HRC.x > .x Then p_min_HRC.x = .x
            If p_min_HRC.y > .y Then p_min_HRC.y = .y
            If p_min_HRC.z > .z Then p_min_HRC.z = .z
        End With

        glTranslated 0, 0, -obj.Bones(BI).length

        jsp = jsp + 1
        joint_stack(jsp) = obj.Bones(BI).joint_i
    Next BI

    While jsp > 0
        glPopMatrix
        jsp = jsp - 1
    Wend
    glPopMatrix
End Sub
Function ComputeHRCDiameter(ByRef obj As HRCSkeleton) As Single
    Dim BI As Integer
    Dim aux_diam As Single
    ComputeHRCDiameter = 0
    With obj
        For BI = 0 To .NumBones - 1
            aux_diam = ComputeHRCBoneDiameter(.Bones(BI))
            'If (aux_diam > ComputeHRCDiameter) Then
            ComputeHRCDiameter = ComputeHRCDiameter + aux_diam
        Next BI
    End With
End Function

Function GetClosestHRCBone(ByRef obj As HRCSkeleton, ByRef Frame As AFrame, ByVal px As Integer, ByVal py As Integer, ByVal DIST As Double) As Integer
    Dim BI As Integer

    Dim min_z As Single
    Dim sbi As Integer
    Dim nBones As Integer

    Dim vp(4) As Long
    Dim P_matrix(16) As Double

    Dim Sel_BUFF() As Long
    ReDim Sel_BUFF(obj.NumBones * 4)

    Dim width As Integer
    Dim height As Integer

    Dim last_joint As String
    Dim joint_stack() As String
    Dim jsp As String
    Dim rot_mat(16) As Double

    ReDim joint_stack(obj.NumBones * 4)
    jsp = 0

    joint_stack(jsp) = obj.Bones(0).joint_f

    glSelectBuffer obj.NumBones * 4, Sel_BUFF(0)
    glInitNames

    glRenderMode GL_SELECT

    glMatrixMode GL_PROJECTION
    glPushMatrix
    glGetDoublev GL_PROJECTION_MATRIX, P_matrix(0)
    glLoadIdentity
    glGetIntegerv GL_VIEWPORT, vp(0)
    width = vp(2)
    height = vp(3)
    gluPickMatrix px - 1, height - py + 1, 3, 3, vp(0)
    'gluPerspective 60, width / height, 0.1, 10000 'max(0.1 - DIST, 0.1), ComputeHRCDiameter(obj) * 2 - DIST
    glMultMatrixd P_matrix(0)

    glMatrixMode GL_MODELVIEW
    glPushMatrix
    With Frame
        glTranslated .RootTranslationX, 0, 0
        glTranslated 0, -.RootTranslationY, 0
        glTranslated 0, 0, .RootTranslationZ

        BuildRotationMatrixWithQuaternions .RootRotationAlpha, .RootRotationBeta, _
            .RootRotationGamma, rot_mat
        glMultMatrixd rot_mat(0)
    End With
    For BI = 0 To obj.NumBones
        glPushName BI
            While Not (obj.Bones(BI).joint_f = joint_stack(jsp)) And jsp > 0
                glPopMatrix
                jsp = jsp - 1
            Wend
            glPushMatrix

            'glRotated Frame.Rotations(bi).Beta, 0#, 1#, 0#
            'glRotated Frame.Rotations(bi).Alpha, 1#, 0#, 0#
            'glRotated Frame.Rotations(bi).Gamma, 0#, 0#, 1#
            BuildRotationMatrixWithQuaternions Frame.Rotations(BI).alpha, Frame.Rotations(BI).Beta, _
                Frame.Rotations(BI).Gamma, rot_mat
            glMultMatrixd rot_mat(0)

            DrawHRCBone obj.Bones(BI), False

            glTranslated 0, 0, -obj.Bones(BI).length

            jsp = jsp + 1
            joint_stack(jsp) = obj.Bones(BI).joint_i
        glPopName
    Next BI

    While jsp > 0
        glPopMatrix
        jsp = jsp - 1
    Wend
    glPopMatrix

    glMatrixMode GL_PROJECTION
    glPopMatrix

    nBones = glRenderMode(GL_RENDER)
    GetClosestHRCBone = -1
    min_z = -1

    For BI = 0 To nBones - 1
        If CompareLongs(min_z, Sel_BUFF(BI * 4 + 1)) Then
            min_z = Sel_BUFF(BI * 4 + 1)
            GetClosestHRCBone = Sel_BUFF(BI * 4 + 3)
        End If
    Next BI
End Function
Function GetClosestHRCBonePiece(ByRef obj As HRCSkeleton, ByRef Frame As AFrame, ByVal b_index As Integer, ByVal px As Integer, ByVal py As Integer, ByVal DIST As Single) As Integer
    Dim BI As Integer
    Dim pi As Integer

    Dim min_z As Single
    Dim sbi As Integer
    Dim nPieces As Integer

    Dim vp(4) As Long
    Dim P_matrix(16) As Double

    Dim last_joint As String
    Dim joint_stack() As String
    Dim jsp As String
    Dim rot_mat(16) As Double

    ReDim joint_stack(obj.NumBones)
    jsp = 0

    With obj.Bones(b_index)
        Dim Sel_BUFF() As Long
        ReDim Sel_BUFF(.NumResources * 4)

        Dim width As Integer
        Dim height As Integer

        glSelectBuffer .NumResources * 4, Sel_BUFF(0)
        glInitNames

        glRenderMode GL_SELECT

        glMatrixMode GL_PROJECTION
        glPushMatrix
        glGetDoublev GL_PROJECTION_MATRIX, P_matrix(0)
        glLoadIdentity
        glGetIntegerv GL_VIEWPORT, vp(0)
        width = vp(2)
        height = vp(3)
        gluPickMatrix px - 1, height - py + 1, 3, 3, vp(0)
        gluPerspective 60, width / height, 0.1, 10000  'max(0.1 - DIST, 0.1), ComputeHRCDiameter(obj) * 2 - DIST

        glMatrixMode GL_MODELVIEW
        glPushMatrix

        glTranslated Frame.RootTranslationX, 0, 0
        glTranslated 0, -Frame.RootTranslationY, 0
        glTranslated 0, 0, Frame.RootTranslationZ

        BuildRotationMatrixWithQuaternions Frame.RootRotationAlpha, Frame.RootRotationBeta, _
            Frame.RootRotationGamma, rot_mat
        glMultMatrixd rot_mat(0)
        For BI = 0 To b_index - 1
            While Not (obj.Bones(BI).joint_f = joint_stack(jsp)) And jsp > 0
                glPopMatrix
                jsp = jsp - 1
            Wend
            glPushMatrix

            'glRotated Frame.Rotations(bi).Beta, 0#, 1#, 0#
            'glRotated Frame.Rotations(bi).Alpha, 1#, 0#, 0#
            'glRotated Frame.Rotations(bi).Gamma, 0#, 0#, 1#
            BuildRotationMatrixWithQuaternions Frame.Rotations(BI).alpha, Frame.Rotations(BI).Beta, _
                Frame.Rotations(BI).Gamma, rot_mat
            glMultMatrixd rot_mat(0)

            glTranslated 0, 0, -obj.Bones(BI).length

            jsp = jsp + 1
            joint_stack(jsp) = obj.Bones(BI).joint_i
        Next BI

        While Not (obj.Bones(b_index).joint_f = joint_stack(jsp)) And jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend
        glPushMatrix

        glRotated Frame.Rotations(b_index).Beta, 0#, 1#, 0#
        glRotated Frame.Rotations(b_index).alpha, 1#, 0#, 0#
        glRotated Frame.Rotations(b_index).Gamma, 0#, 0#, 1#
        jsp = jsp + 1

        For pi = 0 To .NumResources - 1
            glPushName pi
                DrawRSBResource .Resources(pi), False
            glPopName
        Next pi
    End With

    While jsp > 0
        glPopMatrix
        jsp = jsp - 1
    Wend
    glPopMatrix
    glMatrixMode GL_PROJECTION
    glPopMatrix

    nPieces = glRenderMode(GL_RENDER)
    GetClosestHRCBonePiece = -1
    min_z = -1

    For pi = 0 To nPieces - 1
        If CompareLongs(min_z, Sel_BUFF(pi * 4 + 1)) Then
            min_z = Sel_BUFF(pi * 4 + 1)
            GetClosestHRCBonePiece = Sel_BUFF(pi * 4 + 3)
        End If
    Next pi
    ''Debug.Print GetClosestHRCBonePiece, nPieces
End Function
Sub SelectHRCBoneAndPiece(ByRef obj As HRCSkeleton, ByRef Frame As AFrame, ByVal b_index As Integer, ByVal p_index As Integer)
    Dim i As Integer
    Dim jsp As Integer
    Dim rot_mat(16) As Double

    glMatrixMode GL_MODELVIEW
    glPushMatrix
    With Frame
        glTranslated .RootTranslationX, 0, 0
        glTranslated 0, -.RootTranslationY, 0
        glTranslated 0, 0, .RootTranslationZ

        BuildRotationMatrixWithQuaternions .RootRotationAlpha, .RootRotationBeta, _
            .RootRotationGamma, rot_mat
        glMultMatrixd rot_mat(0)
    End With

    If b_index > -1 Then
        jsp = MoveToHRCBone(obj, Frame, b_index)
        DrawHRCBoneBoundingBox obj.Bones(b_index)
        If p_index > -1 Then _
            DrawHRCBonePieceBoundingBox obj.Bones(b_index), p_index

        For i = 0 To jsp
            glPopMatrix
        Next i
    End If
    glPopMatrix
End Sub
Function MoveToHRCBone(ByRef obj As HRCSkeleton, ByRef Frame As AFrame, ByVal b_index As Integer) As Integer
    Dim BI As Integer
    Dim last_joint As String
    Dim joint_stack() As String
    Dim jsp As String

    glMatrixMode GL_MODELVIEW

    ReDim joint_stack(obj.NumBones)
    jsp = 0

    joint_stack(jsp) = obj.Bones(0).joint_f
    For BI = 0 To b_index - 1
        While Not (obj.Bones(BI).joint_f = joint_stack(jsp)) And jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend
        glPushMatrix

        With Frame.Rotations(BI)
            glRotated .Beta, 0#, 1#, 0#
            glRotated .alpha, 1#, 0#, 0#
            glRotated .Gamma, 0#, 0#, 1#
        End With

        glTranslated 0, 0, -obj.Bones(BI).length

        jsp = jsp + 1
        joint_stack(jsp) = obj.Bones(BI).joint_i
    Next BI

    While Not (obj.Bones(b_index).joint_f = joint_stack(jsp)) And jsp > 0
        glPopMatrix
        jsp = jsp - 1
    Wend
    glPushMatrix

    With Frame.Rotations(b_index)
        glRotated .Beta, 0#, 1#, 0#
        glRotated .alpha, 1#, 0#, 0#
        glRotated .Gamma, 0#, 0#, 1#
    End With

    MoveToHRCBone = jsp + 1
End Function
Sub ApplyHRCChanges(ByRef obj As HRCSkeleton, ByRef Frame As AFrame, ByVal merge As Boolean)
    Dim BI As Integer
    Dim last_joint As String
    Dim joint_stack() As String
    Dim jsp As String

    On Error GoTo hand

    glMatrixMode GL_MODELVIEW

    ReDim joint_stack(obj.NumBones)
    jsp = 0

    joint_stack(jsp) = obj.Bones(0).joint_f
    For BI = 0 To obj.NumBones - 1
        While Not (obj.Bones(BI).joint_f = joint_stack(jsp)) And jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend
        glPushMatrix

        glRotated Frame.Rotations(BI).Beta, 0#, 1#, 0#
        glRotated Frame.Rotations(BI).alpha, 1#, 0#, 0#
        glRotated Frame.Rotations(BI).Gamma, 0#, 0#, 1#

        ''Debug.Print bi, obj.Bones(bi).joint_f, obj.Bones(bi).NumResources
        ApplyHRCBoneChanges obj.Bones(BI), ComputeHRCDiameter(obj) * 2, merge

        glTranslated 0, 0, -obj.Bones(BI).length

        jsp = jsp + 1
        joint_stack(jsp) = obj.Bones(BI).joint_i
    Next BI

    While jsp > 0
        glPopMatrix
        jsp = jsp - 1
    Wend
    Exit Sub
hand:
    If Err <> 32755 Then
        MsgBox "Error" + Str(Err), vbOKOnly, "Unknow error ApplyHRCChanges"
    End If
End Sub

Sub GetTwoAFramesInterpolation(ByRef skeleton As HRCSkeleton, ByRef frame_a As AFrame, ByRef frame_b As AFrame, ByVal alpha As Single, ByRef frame_out As AFrame)
    Dim BI As Integer
    Dim last_joint As String
    Dim joint_stack() As String
    Dim rotations_stack_a() As Quaternion
    Dim rotations_stack_b() As Quaternion
    Dim rotations_stack_acum() As Quaternion
    Dim jsp As Integer
    Dim quat_a As Quaternion
    Dim quat_b As Quaternion
    Dim quat_acum_a As Quaternion
    Dim quat_acum_b As Quaternion
    Dim quat_acum_inverse As Quaternion
    Dim quat_interp As Quaternion
    Dim quat_interp_final As Quaternion
    Dim euler_res As Point3D
    Dim alpha_inv As Single
    Dim num_bones As Integer
    Dim mat(16) As Double

    With frame_out
        quat_a = GetQuaternionFromEulerYXZr(frame_a.RootRotationAlpha, frame_a.RootRotationBeta, frame_a.RootRotationGamma)
        quat_b = GetQuaternionFromEulerYXZr(frame_b.RootRotationAlpha, frame_b.RootRotationBeta, frame_b.RootRotationGamma)
        quat_interp = QuaternionSlerp2(quat_a, quat_b, alpha)
        BuildMatrixFromQuaternion quat_interp, mat
        euler_res = GetEulerYXZrFromMatrix(mat)

        .RootRotationAlpha = euler_res.y
        .RootRotationBeta = euler_res.x
        .RootRotationGamma = euler_res.z

        alpha_inv = 1# - alpha
        .RootTranslationX = frame_a.RootTranslationX * alpha_inv + frame_b.RootTranslationX * alpha
        .RootTranslationY = frame_a.RootTranslationY * alpha_inv + frame_b.RootTranslationY * alpha
        .RootTranslationZ = frame_a.RootTranslationZ * alpha_inv + frame_b.RootTranslationZ * alpha

        num_bones = skeleton.NumBones + 1

        ReDim joint_stack(num_bones)
        ReDim rotations_stack_a(num_bones)
        ReDim rotations_stack_b(num_bones)
        ReDim rotations_stack_acum(num_bones)
        ReDim .Rotations(skeleton.NumBones)
        jsp = 1

        rotations_stack_a(0) = quat_a
        rotations_stack_b(0) = quat_b
        rotations_stack_acum(0) = quat_interp

        joint_stack(jsp) = skeleton.Bones(0).joint_f
        For BI = 0 To skeleton.NumBones - 1
            While Not (skeleton.Bones(BI).joint_f = joint_stack(jsp)) And jsp > 0
                jsp = jsp - 1
            Wend

            quat_a = GetQuaternionFromEulerYXZr(frame_a.Rotations(BI).alpha, frame_a.Rotations(BI).Beta, frame_a.Rotations(BI).Gamma)
            MultiplyQuaternions rotations_stack_a(jsp - 1), quat_a, quat_acum_a
            NormalizeQuaternion quat_acum_a
            rotations_stack_a(jsp) = quat_acum_a
            quat_b = GetQuaternionFromEulerYXZr(frame_b.Rotations(BI).alpha, frame_b.Rotations(BI).Beta, frame_b.Rotations(BI).Gamma)
            MultiplyQuaternions rotations_stack_b(jsp - 1), quat_b, quat_acum_b
            NormalizeQuaternion quat_acum_b
            rotations_stack_b(jsp) = quat_acum_b

            quat_interp = QuaternionSlerp2(quat_acum_a, quat_acum_b, alpha)
            rotations_stack_acum(jsp) = quat_interp
            quat_acum_inverse = GetQuaternionConjugate(rotations_stack_acum(jsp - 1))
            MultiplyQuaternions quat_acum_inverse, quat_interp, quat_interp_final
            NormalizeQuaternion quat_interp_final
            BuildMatrixFromQuaternion quat_interp_final, mat
            euler_res = GetEulerYXZrFromMatrix(mat)

            .Rotations(BI).alpha = euler_res.y
            .Rotations(BI).Beta = euler_res.x
            .Rotations(BI).Gamma = euler_res.z

            jsp = jsp + 1
            joint_stack(jsp) = skeleton.Bones(BI).joint_i
        Next BI
    End With
End Sub
Sub InterpolateFramesAAnimation(ByRef hrc_sk As HRCSkeleton, ByRef obj As AAnimation, ByVal base_frame As Integer, ByVal num_interpolated_frames As Integer)
    Dim fi As Integer
    Dim alpha As Single

    With obj
        'Create new frames
        .NumFrames = .NumFrames + num_interpolated_frames
        ReDim Preserve .Frames(.NumFrames - 1)
        'Move the original frames into their new positions
        For fi = .NumFrames - 1 To base_frame + num_interpolated_frames Step -1
            .Frames(fi) = .Frames(fi - num_interpolated_frames)
        Next fi

        'Interpolate the new frames
        For fi = 1 To num_interpolated_frames
            alpha = CSng(fi) / CSng(num_interpolated_frames + 1)
            GetTwoAFramesInterpolation hrc_sk, .Frames(base_frame), .Frames(base_frame + num_interpolated_frames + 1), alpha, .Frames(base_frame + fi)
        Next fi
    End With
End Sub
Sub InterpolateAAnimation(ByRef hrc_sk As HRCSkeleton, ByRef obj As AAnimation, ByVal num_interpolated_frames As Integer, ByVal is_loop As Boolean)
    Dim alpha As Single
    Dim fi As Integer
    Dim ifi As Integer
    Dim next_elem_diff As Integer
    Dim frame_offset As Integer
    Dim base_final_frame As Integer

    next_elem_diff = num_interpolated_frames + 1

    frame_offset = 0
    If Not is_loop Then
        frame_offset = num_interpolated_frames
    End If

    With obj
        If .NumFrames = 1 Then
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

        If is_loop Then
            base_final_frame = .NumFrames - num_interpolated_frames - 1
            For ifi = 1 To num_interpolated_frames
                alpha = CSng(ifi) / CSng(num_interpolated_frames + 1)
                GetTwoAFramesInterpolation hrc_sk, .Frames(base_final_frame), .Frames(0), alpha, .Frames(base_final_frame + ifi)
            Next ifi
        End If
    End With
End Sub

Sub FixAAnimation(ByRef hrc_sk As HRCSkeleton, ByRef obj As AAnimation)
    Dim fi As Integer
    Dim base_fi As Integer
    Dim num_broken_frames As Integer

    With obj
        While fi <= .NumFrames - 1
            base_fi = fi
            num_broken_frames = 0
            While IsFrameBrokenAAnimation(obj, fi)
                RemoveFrameAAnimation obj, fi
                num_broken_frames = num_broken_frames + 1
            Wend

            If num_broken_frames > 0 Then
                If fi = 0 Then
                    fi = fi + 1
                Else
                    InterpolateFramesAAnimation hrc_sk, obj, fi - 1, num_broken_frames
                    fi = fi + num_broken_frames
                End If
            Else
                fi = fi + 1
            End If

        Wend
    End With
End Sub
