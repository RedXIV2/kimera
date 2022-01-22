Attribute VB_Name = "FF7AASkeleton"
Option Explicit
Type AASkeleton
    filename As String
    unk(3) As Long
    NumBones As Long
    unk2(2) As Long
    NumTextures As Long
    NumBodyAnims As Long
    unk3(2) As Long
    NumWeaponAnims As Long
    unk4(2) As Long
    Bones() As AABone
    textures() As TEXTexture
    NumWeapons As Integer
    WeaponModels() As PModel
    TexIDS() As Long
    IsBattleLocation As Boolean
    IsLimitBreak As Boolean
End Type
Sub ReadAASkeleton(ByVal filename As String, ByRef skeleton As AASkeleton, ByVal is_limit_breakQ As Boolean, ByVal load_geometryQ As Boolean)
    Dim fileNumber As Integer
    Dim pSufix1 As Integer
    Dim pSufix2 As Integer
    Dim baseName As String
    Dim weaponFileName As String
    Dim texFileName As String
    Dim BI As Integer
    Dim ti As Integer
    Dim B As Boolean
    Dim pSuffix2End As Integer
    Dim fixfield As Integer

    On Error GoTo ErrHandRead

    ChDir GetPathFromString(filename)

    fileNumber = FreeFile
    Open filename For Binary As fileNumber

    With skeleton
        .filename = TrimPath(filename)
        Get fileNumber, 1, .unk
        Get fileNumber, 13, .NumBones
        Get fileNumber, 17, .unk2
        Get fileNumber, 25, .NumTextures
        Get fileNumber, 29, .NumBodyAnims
        Get fileNumber, 33, .unk3
        Get fileNumber, 41, .NumWeaponAnims
        Get fileNumber, 45, .unk4

        'If is_limit_breakQ Then
        '    .NumBodyAnims = 8
        '    .NumWeaponAnims = 8
        'End If

        .IsLimitBreak = is_limit_breakQ
        baseName = Left$(filename, Len(filename) - 2)
        pSufix1 = 97
        B = False

        If .NumBones = 0 Then   'It's a battle location model
            .IsBattleLocation = True

            For pSufix1 = 97 To 123

                If pSufix1 = 97 Then fixfield = 109 Else fixfield = 97

                For pSufix2 = fixfield To 123
                    If FileExist(baseName + Chr$(pSufix1) + Chr$(pSufix2)) Then
                        ReDim Preserve .Bones(.NumBones)
                        If load_geometryQ Then
                            ReadAABattleLocationPiece.Bones(.NumBones), .NumBones, baseName + Chr$(pSufix1) + Chr$(pSufix2)
                        End If
                        .NumBones = .NumBones + 1
                    End If
                Next pSufix2

            Next pSufix1

            pSufix1 = 97
        Else                    'It's a character battle model
            .IsBattleLocation = False
            pSufix2 = 109
            ReDim .Bones(.NumBones - 1)
            For BI = 0 To .NumBones - 1
                ReadAABone fileNumber, 53 + BI * 12, baseName + Chr$(pSufix1) + Chr$(pSufix2), .Bones(BI), load_geometryQ
                If pSufix2 >= 122 Then
                    pSufix1 = pSufix1 + 1
                    pSufix2 = 97
                Else
                    pSufix2 = pSufix2 + 1
                End If
            Next BI

            'Read weapon models
            pSufix1 = 99
            .NumWeapons = 0
            ReDim .WeaponModels(122 - 107)
            For pSufix2 = 107 To 122
                weaponFileName = baseName + Chr$(pSufix1) + Chr$(pSufix2)
                If FileExist(weaponFileName) Then
                    If load_geometryQ Then
                        ReadPModel .WeaponModels(.NumWeapons), weaponFileName
                    End If
                    ''Debug.Print "Loaded weapon model " + weaponFileName
                    .NumWeapons = .NumWeapons + 1
                End If
            Next pSufix2
            ReDim Preserve .WeaponModels(.NumWeapons)
        End If

        'Read Textures
        pSufix1 = 97
        pSufix2 = 99

        If load_geometryQ Then
            ReDim .TexIDS(.NumTextures)
            ReDim .textures(.NumTextures)
            ti = 0
            pSuffix2End = 99 + .NumTextures - 1
            For pSufix2 = 99 To pSuffix2End
                texFileName = baseName + Chr$(pSufix1) + Chr$(pSufix2)
                'If FileExist(texFileName) Then
                    .textures(ti).tex_file = texFileName
                    If ReadTEXTexture(.textures(ti), texFileName) = 0 Then
                        LoadTEXTexture .textures(ti)
                        LoadBitmapFromTEXTexture .textures(ti)
                    End If
                    .TexIDS(ti) = .textures(ti).tex_id
                'End If
                ti = ti + 1
            Next pSufix2
        End If
    End With

    Close fileNumber
    Exit Sub
ErrHandRead:
    'Debug.Print "Error reading AA file!!!"
    MsgBox "Error reading AA file " + filename + "!!!", vbOKOnly, "Error reading"
End Sub
Sub ReadMagicSkeleton(ByVal filename As String, ByRef skeleton As AASkeleton, ByVal load_geometryQ As Boolean)
    Dim fileNumber As Integer
    Dim pSufix As String
    Dim tSufix As String
    Dim baseName As String
    Dim texFileName As String
    Dim BI As Integer
    Dim zi As Integer
    Dim ti As Integer

    On Error GoTo ErrHandRead

    ChDir GetPathFromString(filename)

    fileNumber = FreeFile
    Open filename For Binary As fileNumber

    With skeleton
        .filename = TrimPath(filename)
        Get fileNumber, 1, .unk
        Get fileNumber, 13, .NumBones
        Get fileNumber, 17, .unk2
        Get fileNumber, 25, .NumTextures
        Get fileNumber, 29, .NumBodyAnims
        Get fileNumber, 33, .unk3
        Get fileNumber, 41, .NumWeaponAnims
        Get fileNumber, 45, .unk4

        baseName = Left$(filename, Len(filename) - 1)

        .IsBattleLocation = False
        .IsLimitBreak = False
        ReDim .Bones(.NumBones - 1)
        For BI = 0 To .NumBones - 1
            pSufix = "p"
            For zi = 0 To 2 - Len(Str$(BI))
                pSufix = pSufix + Right$(Str$(0), 1)
            Next zi
            pSufix = pSufix + Right$(Str$(BI), Len(Str$(BI)) - 1)
            ReadAABone fileNumber, 53 + BI * 12, baseName + pSufix, .Bones(BI), load_geometryQ
        Next BI

        If load_geometryQ Then
            ReDim .TexIDS(0)
            .NumTextures = 0
            For ti = 0 To 100
                tSufix = "t"
                For zi = 0 To 2 - Len(Str$(ti))
                    tSufix = tSufix + Right$(Str$(0), 1)
                Next zi
                tSufix = tSufix + Right$(Str$(ti), Len(Str$(ti)) - 1)

                texFileName = baseName + tSufix
                If FileExist(texFileName) Then
                    ReDim Preserve .textures(.NumTextures)
                    ReDim Preserve .TexIDS(.NumTextures)
                    .textures(.NumTextures).tex_file = texFileName
                    If ReadTEXTexture(.textures(.NumTextures), texFileName) = 0 Then
                        LoadTEXTexture .textures(.NumTextures)
                        LoadBitmapFromTEXTexture .textures(.NumTextures)
                    End If
                    .TexIDS(.NumTextures) = .textures(.NumTextures).tex_id
                    .NumTextures = .NumTextures + 1
                End If
            Next ti
        End If
    End With

    Close fileNumber
    Exit Sub
ErrHandRead:
    'Debug.Print "Error reading D file!!!"
    MsgBox "Error reading D file " + filename + "!!!", vbOKOnly, "Error reading"
End Sub

Sub WriteAASkeleton(ByVal filename As String, ByRef skeleton As AASkeleton)
    Dim fileNumber As Integer
    Dim pSufix1 As Integer
    Dim pSufix2 As Integer
    Dim baseName As String
    Dim BI As Integer
    Dim wi As Integer
    Dim ti As Integer
    Dim aux_bones_battle_stance As Integer
    aux_bones_battle_stance = 0

    On Error GoTo ErrHandWrite

    ChDir GetPathFromString(filename)

    fileNumber = FreeFile
    Open filename For Output As fileNumber
    Close fileNumber
    Open filename For Binary As fileNumber

    With skeleton
        baseName = Left$(filename, Len(filename) - 2)
        pSufix1 = 97
        pSufix2 = 109
        For BI = 0 To .NumBones - 1
            WriteAABone fileNumber, 53 + BI * 12, baseName + Chr$(pSufix1) + Chr$(pSufix2), .Bones(BI)
            If pSufix2 >= 122 Then
                pSufix1 = pSufix1 + 1
                pSufix2 = 97
            Else
                pSufix2 = pSufix2 + 1
            End If
        Next BI

        Put fileNumber, 1, .unk
        If .IsBattleLocation Then
            Put fileNumber, 13, aux_bones_battle_stance
        Else
            Put fileNumber, 13, .NumBones
        End If
        Put fileNumber, 17, .unk2
        Put fileNumber, 25, .NumTextures
        Put fileNumber, 29, .NumBodyAnims
        Put fileNumber, 33, .unk3
        Put fileNumber, 41, .NumWeaponAnims
        Put fileNumber, 45, .unk4

        pSufix1 = 99
        pSufix2 = 107
        For wi = 0 To .NumWeapons - 1
            WritePModel .WeaponModels(wi), baseName + Chr$(pSufix1) + Chr$(pSufix2 + wi)
        Next wi

        pSufix1 = 97
        pSufix2 = 99
        For ti = 0 To .NumTextures - 1
            WriteTEXTexture .textures(ti), baseName + Chr$(pSufix1) + Chr$(pSufix2 + ti)
        Next ti
    End With

    Close fileNumber
    Exit Sub
ErrHandWrite:
    'Debug.Print "Error writting AA file!!!"
    MsgBox "Error writting AA file " + filename + "!!!", vbOKOnly, "Error writting"
End Sub
Sub WriteMagicSkeleton(ByVal filename As String, ByRef skeleton As AASkeleton)
    Dim fileNumber As Integer
    Dim pSufix As String
    Dim tSufix As String
    Dim baseName As String
    Dim texFileName As String
    Dim BI As Integer
    Dim zi As Integer
    Dim ti As Integer

    On Error GoTo ErrHandWrite

    ChDir GetPathFromString(filename)

    fileNumber = FreeFile
    Open filename For Output As fileNumber
    Close fileNumber
    Open filename For Binary As fileNumber

    With skeleton
        baseName = Left$(filename, Len(filename) - 1)

        'ReDim .Bones(.NumBones)
        For BI = 0 To .NumBones - 1
            pSufix = "p"
            For zi = 0 To 2 - Len(Str$(BI))
                pSufix = pSufix + Right$(Str$(0), 1)
            Next zi
            pSufix = pSufix + Right$(Str$(BI), Len(Str$(BI)) - 1)
            WriteAABone fileNumber, 53 + BI * 12, baseName + pSufix, .Bones(BI)
        Next BI

        Put fileNumber, 1, .unk
        Put fileNumber, 13, .NumBones
        Put fileNumber, 17, .unk2
        Put fileNumber, 25, .NumTextures
        Put fileNumber, 29, .NumBodyAnims
        Put fileNumber, 33, .unk3
        Put fileNumber, 41, .NumWeaponAnims
        Put fileNumber, 45, .unk4

        For ti = 0 To .NumTextures - 1
            pSufix = "t"
            For zi = 0 To 2 - Len(Str$(ti))
                pSufix = pSufix + Right$(Str$(0), 1)
            Next zi
            pSufix = pSufix + Right$(Str$(ti), Len(Str$(ti)) - 1)
            WriteTEXTexture .textures(ti), baseName + pSufix '.textures(ti).tex_file
        Next ti
    End With

    Close fileNumber
    Exit Sub
ErrHandWrite:
    'Debug.Print "Error writting D file!!!"
    MsgBox "Error writting D file " + filename + "!!!", vbOKOnly, "Error writting"
End Sub
Sub CreateDListsFromAASkeleton(ByRef obj As AASkeleton)
    Dim BI As Integer

    With obj
        For BI = 0 To .NumBones - 1
            CreateDListsFromAASkeletonBone .Bones(BI)
        Next BI
    End With
End Sub
Sub FreeAASkeletonResources(ByRef obj As AASkeleton)
    Dim BI As Integer
    Dim ti As Integer
    Dim wi As Integer

    With obj
        For BI = 0 To .NumBones - 1
            FreeAABoneResources .Bones(BI)
        Next BI
    End With

    For ti = 0 To obj.NumTextures - 1
        With obj.textures(ti)
            glDeleteTextures 1, .tex_id
            DeleteDC .hdc
            DeleteObject .hbmp
        End With
    Next ti

    With obj
        For wi = 0 To .NumWeapons - 1
            FreePModelResources .WeaponModels(wi)
        Next wi
    End With

End Sub
Sub DrawAASkeleton(ByRef obj As AASkeleton, ByRef Frame As DAFrame, ByRef FrameWeapon As DAFrame, ByVal WeaponId As Integer, ByVal UseDLists As Boolean)
    Dim BI As Integer
    Dim joint_stack() As Integer
    Dim jsp As Integer
    Dim rot_mat(16) As Double

    glMatrixMode GL_MODELVIEW

    ReDim joint_stack(obj.NumBones)
    jsp = 0

    glPushMatrix
    glTranslated Frame.X_start, Frame.Y_start, Frame.Z_start
    With Frame.Bones(0)
        'Debug.Print .alpha; ", "; .Beta; ", "; .Gamma
        BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
    End With
    glMultMatrixd rot_mat(0)

    joint_stack(jsp) = -1
    For BI = 0 To obj.NumBones - 1
        If obj.IsBattleLocation Then
            DrawAASkeletonBone obj.Bones(BI), obj.TexIDS, False
        Else
            While Not (obj.Bones(BI).ParentBone = joint_stack(jsp)) And jsp > 0
                glPopMatrix
                jsp = jsp - 1
            Wend

            glPushMatrix

            'glRotated Frame.Bones(bi + 1).Beta, 0#, 1#, 0#
            'glRotated Frame.Bones(bi + 1).alpha, 1#, 0#, 0#
            'glRotated Frame.Bones(bi + 1).Gamma, 0#, 0#, 1#
            With Frame.Bones(BI + IIf(obj.NumBones > 1, 1, 0))
                BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
            End With
            glMultMatrixd rot_mat(0)


            DrawAASkeletonBone obj.Bones(BI), obj.TexIDS, UseDLists

            glTranslated 0, 0, obj.Bones(BI).length

            jsp = jsp + 1
            joint_stack(jsp) = BI
        End If
    Next BI

    If Not obj.IsBattleLocation Then
        While jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend
    End If
    glPopMatrix

    If (WeaponId > -1 And obj.NumWeapons > 0) Then
        glPushMatrix
        glTranslated FrameWeapon.X_start, FrameWeapon.Y_start, FrameWeapon.Z_start
        'glRotated FrameWeapon.Bones(0).Beta, 0#, 1#, 0#
        'glRotated FrameWeapon.Bones(0).Alpha, 1#, 0#, 0#
        'glRotated FrameWeapon.Bones(0).Gamma, 0#, 0#, 1#
        With FrameWeapon.Bones(0)
            BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
        End With
        glMultMatrixd rot_mat(0)

        glMatrixMode GL_MODELVIEW
        glPushMatrix
        With obj.WeaponModels(WeaponId)
            glTranslatef .RepositionX, .RepositionY, .RepositionZ

            glRotated .RotateAlpha, 1#, 0#, 0#
            glRotated .RotateBeta, 0#, 1#, 0#
            glRotated .RotateGamma, 0#, 0#, 1#

            glScalef .ResizeX, .ResizeY, .ResizeZ
        End With
        If UseDLists Then
            DrawPModelDLists obj.WeaponModels(WeaponId), obj.TexIDS
        Else
            DrawPModel obj.WeaponModels(WeaponId), obj.TexIDS, False
        End If
        glPopMatrix

        glPopMatrix
    End If
End Sub
Sub DrawAASkeletonBones(ByRef obj As AASkeleton, ByRef Frame As DAFrame)
    Dim BI As Integer

    Dim joint_stack() As Integer
    Dim jsp As Integer
    Dim rot_mat(16) As Double

    If obj.IsBattleLocation Then Exit Sub

    glMatrixMode GL_MODELVIEW

    ReDim joint_stack(obj.NumBones)
    jsp = 0

    joint_stack(jsp) = -1

    glPointSize 5


    glPushMatrix
    glTranslated Frame.X_start, Frame.Y_start, Frame.Z_start
    With Frame.Bones(0)
        BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
    End With
    glMultMatrixd rot_mat(0)

    For BI = 0 To obj.NumBones - 1
        While Not (obj.Bones(BI).ParentBone = joint_stack(jsp)) And jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend
        glPushMatrix


        'glRotated Frame.Bones(bi + 1).Beta, 0#, 1#, 0#
        'glRotated Frame.Bones(bi + 1).Alpha, 1#, 0#, 0#
        'glRotated Frame.Bones(bi + 1).Gamma, 0#, 0#, 1#
        With Frame.Bones(BI + IIf(obj.NumBones > 1, 1, 0))
            BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
        End With
        glMultMatrixd rot_mat(0)

        glBegin GL_POINTS
            glVertex3f 0, 0, 0
            glVertex3f 0, 0, obj.Bones(BI).length
        glEnd

        glBegin GL_LINES
            glVertex3f 0, 0, 0
            glVertex3f 0, 0, obj.Bones(BI).length
        glEnd

        glTranslated 0, 0, obj.Bones(BI).length

        jsp = jsp + 1
        joint_stack(jsp) = BI
    Next BI

    If Not obj.IsBattleLocation Then
        While jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend
    End If
    glPopMatrix
End Sub
Sub SetCameraAASkeleton(ByRef obj As AASkeleton, ByVal cx As Single, ByVal cy As Single, ByVal CZ As Single, ByVal alpha As Single, ByVal Beta As Single, ByVal Gamma As Single, ByVal redX As Single, ByVal redY As Single, ByVal redZ As Single)
    Dim width As Integer
    Dim height As Integer
    Dim vp(4) As Long

    glGetIntegerv GL_VIEWPORT, vp(0)
    width = vp(2)
    height = vp(3)

    glMatrixMode GL_PROJECTION
    glLoadIdentity
    'gluPerspective 60, width / height, max(0.1 - CZ, 0.1), max(100000 - CZ, 0.1) 'max(0.1 - CZ, 0.1),ComputeAADiameter(obj) * 4 - CZ
    gluPerspective 60, width / height, 0.1, 10000

    Dim f_start As Single
    Dim f_end As Single
    f_start = 500 - CZ
    f_end = 100000 - CZ
    glFogfv GL_FOG_START, f_start
    glFogfv GL_FOG_END, f_end

    glMatrixMode GL_MODELVIEW
    glLoadIdentity

    glTranslatef cx, cy, CZ - ComputeAADiameter(obj) * 2

    glRotatef Beta, 0#, 1#, 0#
    glRotatef alpha, 1#, 0#, 0#
    glRotatef Gamma, 0#, 0#, 1#

    glScalef redX, redY, redZ
End Sub
Sub ComputeAABoundingBox(ByRef obj As AASkeleton, ByRef Frame As DAFrame, ByRef p_min_AA As Point3D, ByRef p_max_AA As Point3D)
    Dim last_joint As String
    Dim joint_stack() As String
    Dim jsp As Integer

    ReDim joint_stack(obj.NumBones * 4)
    jsp = 0
    joint_stack(jsp) = -1

    Dim rot_mat(16) As Double
    Dim MV_matrix(16) As Double
    Dim BI As Integer

    Dim p_max_bone As Point3D
    Dim p_min_bone As Point3D

    Dim p_max_bone_trans As Point3D
    Dim p_min_bone_trans As Point3D

    p_max_AA.x = -INFINITY_SINGLE
    p_max_AA.y = -INFINITY_SINGLE
    p_max_AA.z = -INFINITY_SINGLE

    p_min_AA.x = INFINITY_SINGLE
    p_min_AA.y = INFINITY_SINGLE
    p_min_AA.z = INFINITY_SINGLE

    glMatrixMode GL_MODELVIEW
    glPushMatrix
    glLoadIdentity
    glTranslated Frame.X_start, Frame.Y_start, Frame.Z_start
    With Frame.Bones(0)
        BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
    End With
    glMultMatrixd rot_mat(0)
    For BI = 0 To obj.NumBones - 1
        While Not (obj.Bones(BI).ParentBone = joint_stack(jsp)) And jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend
        glPushMatrix

        With Frame.Bones(BI + IIf(obj.NumBones > 1, 1, 0))
            BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
        End With
        glMultMatrixd rot_mat(0)

        ComputeAABoneBoundingBox obj.Bones(BI), p_min_bone, p_max_bone

        glGetDoublev GL_MODELVIEW_MATRIX, MV_matrix(0)

        ComputeTransformedBoxBoundingBox MV_matrix, p_min_bone, p_max_bone, _
            p_min_bone_trans, p_max_bone_trans

        With p_max_bone_trans
            If p_max_AA.x < .x Then p_max_AA.x = .x
            If p_max_AA.y < .y Then p_max_AA.y = .y
            If p_max_AA.z < .z Then p_max_AA.z = .z
        End With

        With p_min_bone_trans
            If p_min_AA.x > .x Then p_min_AA.x = .x
            If p_min_AA.y > .y Then p_min_AA.y = .y
            If p_min_AA.z > .z Then p_min_AA.z = .z
        End With

        glTranslated 0, 0, obj.Bones(BI).length

        jsp = jsp + 1
        joint_stack(jsp) = BI
    Next BI

    While jsp > 0
        glPopMatrix
        jsp = jsp - 1
    Wend
    glPopMatrix
End Sub
Function ComputeAADiameter(ByRef obj As AASkeleton) As Single
    Dim BI As Integer

    Dim MaxPath As Long
    Dim currentPath As Long

    Dim joint_stack() As Integer
    Dim jsp As Integer

    ReDim joint_stack(obj.NumBones)
    jsp = 0

    With obj
        If obj.IsBattleLocation Then
            For BI = 0 To .NumBones - 2
                If .Bones(BI).length > MaxPath Then MaxPath = .Bones(BI).length
            Next BI
        Else
            joint_stack(jsp) = -1

            For BI = 0 To .NumBones - 1
                While Not (.Bones(BI).ParentBone = joint_stack(jsp)) And jsp > 0
                    currentPath = currentPath + .Bones(joint_stack(jsp)).length
                    jsp = jsp - 1
                Wend
                currentPath = currentPath - .Bones(BI).length
                If currentPath > MaxPath Then MaxPath = currentPath
                jsp = jsp + 1
                joint_stack(jsp) = BI
            Next BI
        End If
    End With

    ComputeAADiameter = MaxPath
End Function

Function GetClosestAABone(ByRef obj As AASkeleton, ByRef Frame As DAFrame, ByRef FrameWeapon As DAFrame, ByVal WeaponId As Integer, ByVal px As Integer, ByVal py As Integer, ByVal DIST As Single) As Integer
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

    Dim joint_stack() As Integer
    Dim jsp As Integer
    Dim textures(0) As Long
    Dim rot_mat(16) As Double

    ReDim joint_stack(obj.NumBones * 4)
    jsp = 0

    joint_stack(jsp) = -1

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
   ' gluPerspective 60, width / height, 0.1, 10000  'max(0.1 - DIST, 0.1), ComputeAADiameter(obj) * 4 - DIST
    glMultMatrixd P_matrix(0)
    glMatrixMode GL_MODELVIEW

    glPushMatrix
    glTranslated Frame.X_start, Frame.Y_start, Frame.Z_start
    With Frame.Bones(0)
        BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
    End With
    glMultMatrixd rot_mat(0)
    For BI = 0 To obj.NumBones - 1
        glPushName BI
            If obj.IsBattleLocation Then
                DrawAASkeletonBone obj.Bones(BI), obj.TexIDS, False
            Else
                While Not (obj.Bones(BI).ParentBone = joint_stack(jsp)) And jsp > 0
                    glPopMatrix
                    jsp = jsp - 1
                Wend
                glPushMatrix

                'glRotated Frame.Bones(bi + 1).Beta, 0#, 1#, 0#
                'glRotated Frame.Bones(bi + 1).Alpha, 1#, 0#, 0#
                'glRotated Frame.Bones(bi + 1).Gamma, 0#, 0#, 1#
                With Frame.Bones(BI + IIf(obj.NumBones > 1, 1, 0))
                    BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
                End With
                glMultMatrixd rot_mat(0)

                DrawAASkeletonBone obj.Bones(BI), obj.TexIDS, False

                glTranslated 0, 0, obj.Bones(BI).length

                jsp = jsp + 1
                joint_stack(jsp) = BI
            End If
        glPopName
    Next BI

    If Not obj.IsBattleLocation Then
        While jsp >= 0
            glPopMatrix
            jsp = jsp - 1
        Wend
    End If
    glPopMatrix

    If (WeaponId > -1 And obj.NumWeapons > 0) Then
        glPushMatrix
        glTranslated FrameWeapon.X_start, FrameWeapon.Y_start, FrameWeapon.Z_start
        'glRotated FrameWeapon.Bones(0).Beta, 0#, 1#, 0#
        'glRotated FrameWeapon.Bones(0).Alpha, 1#, 0#, 0#
        'glRotated FrameWeapon.Bones(0).Gamma, 0#, 0#, 1#
        With FrameWeapon.Bones(0)
            BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
        End With
        glMultMatrixd rot_mat(0)

        glPushMatrix
        With obj.WeaponModels(WeaponId)
            glTranslatef .RepositionX, .RepositionY, .RepositionZ

            glRotated .RotateBeta, 0#, 1#, 0#
            glRotated .RotateAlpha, 1#, 0#, 0#
            glRotated .RotateGamma, 0#, 0#, 1#

            glScalef .ResizeX, .ResizeY, .ResizeZ
        End With

        glPushName obj.NumBones
            DrawPModel obj.WeaponModels(WeaponId), obj.TexIDS, False
        glPopName
        glPopMatrix

        glPopMatrix
    End If

    glMatrixMode GL_PROJECTION
    glPopMatrix

    nBones = glRenderMode(GL_RENDER)
    GetClosestAABone = -1
    min_z = -1

    For BI = 0 To nBones - 1
        If CompareLongs(min_z, Sel_BUFF(BI * 4 + 1)) Then
            min_z = Sel_BUFF(BI * 4 + 1)
            GetClosestAABone = Sel_BUFF(BI * 4 + 3)
        End If
    Next BI
End Function
Function GetClosestAABoneModel(ByRef obj As AASkeleton, ByRef Frame As DAFrame, ByVal b_index As Integer, ByVal px As Integer, ByVal py As Integer, ByVal DIST As Single) As Integer
    Dim i As Integer

    Dim BI As Integer
    Dim mi As Integer

    Dim min_z As Single
    Dim sbi As Integer
    Dim nModels As Integer

    Dim tex_ids(0) As Long

    Dim vp(4) As Long
    Dim P_matrix(16) As Double

    Dim jsp As Integer

    With obj.Bones(b_index)
        Dim Sel_BUFF() As Long
        ReDim Sel_BUFF(.NumModels * 4)

        Dim width As Integer
        Dim height As Integer

        glSelectBuffer .NumModels * 4, Sel_BUFF(0)
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
        'gluPerspective 60, width / height, 0.1, 10000  'max(0.1 - DIST, 0.1), ComputeAADiameter(obj) * 4 - DIST
        glMultMatrixd P_matrix(0)

        jsp = MoveToAABone(obj, Frame, b_index)

        For mi = 0 To .NumModels - 1
            glMatrixMode GL_MODELVIEW
            glPushMatrix
            glTranslatef .Models(mi).RepositionX, .Models(mi).RepositionY, _
                .Models(mi).RepositionZ

            glRotated .Models(mi).RotateAlpha, 1#, 0#, 0#
            glRotated .Models(mi).RotateBeta, 0#, 1#, 0#
            glRotated .Models(mi).RotateGamma, 0#, 0#, 1#

            glScalef .Models(mi).ResizeX, .Models(mi).ResizeY, .Models(mi).ResizeZ

            glPushName mi
                DrawPModel .Models(mi), tex_ids(), False
            glPopName
            glPopMatrix
        Next mi

        For i = 0 To jsp
            glPopMatrix
        Next i
        glPopMatrix
        glMatrixMode GL_PROJECTION
        glPopMatrix
    End With

    nModels = glRenderMode(GL_RENDER)
    GetClosestAABoneModel = -1
    min_z = -1

    For mi = 0 To nModels - 1
        If CompareLongs(min_z, Sel_BUFF(mi * 4 + 1)) Then
            min_z = Sel_BUFF(mi * 4 + 1)
            GetClosestAABoneModel = Sel_BUFF(mi * 4 + 3)
        End If
    Next mi
    ''Debug.Print GetClosestAABoneModel, nModels
End Function
Sub SelectAABoneAndModel(ByRef obj As AASkeleton, ByRef Frame As DAFrame, ByRef FrameWeapon As DAFrame, ByVal WeaponId As Integer, ByVal b_index As Integer, ByVal p_index As Integer)
    Dim i As Integer
    Dim jsp As Integer

    If b_index > -1 And b_index < obj.NumBones Then
        jsp = MoveToAABone(obj, Frame, b_index)
        DrawAABoneBoundingBox obj.Bones(b_index)
        If p_index > -1 Then _
            DrawAABoneModelBoundingBox obj.Bones(b_index), p_index

        For i = 0 To jsp
            glPopMatrix
        Next i
    ElseIf b_index = obj.NumBones Then
        DrawAAWeaponBoundingBox obj, FrameWeapon, WeaponId
    End If
End Sub
Sub DrawAAWeaponBoundingBox(ByRef obj As AASkeleton, ByRef Frame As DAFrame, ByVal WeaponId As Integer)
    Dim rot_mat(16) As Double
    If (WeaponId > -1 And obj.NumWeapons > 0) Then
        glPushMatrix
        With Frame
            glTranslated .X_start, .Y_start, .Z_start
            'glRotated .Bones(0).Beta, 0#, 1#, 0#
            'glRotated .Bones(0).Alpha, 1#, 0#, 0#
            'glRotated .Bones(0).Gamma, 0#, 0#, 1#
        End With
        With Frame.Bones(0)
            BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
        End With
        glMultMatrixd rot_mat(0)

        glPushMatrix
        With obj.WeaponModels(WeaponId)
            glTranslatef .RepositionX, .RepositionY, .RepositionZ

            glRotated .RotateBeta, 0#, 1#, 0#
            glRotated .RotateAlpha, 1#, 0#, 0#
            glRotated .RotateGamma, 0#, 0#, 1#

            glScalef .ResizeX, .ResizeY, .ResizeZ
        End With

        DrawPModelBoundingBox obj.WeaponModels(WeaponId)
        glPopMatrix

        glPopMatrix
    End If
End Sub
Function MoveToAABoneMiddle(ByRef obj As AASkeleton, ByRef Frame As DAFrame, ByVal b_index As Integer) As Integer
    MoveToAABoneMiddle = MoveToAABone(obj, Frame, b_index)
    glTranslated 0, 0, obj.Bones(b_index).length / 2
End Function

Function MoveToAABoneEnd(ByRef obj As AASkeleton, ByRef Frame As DAFrame, ByVal b_index As Integer) As Integer
    MoveToAABoneEnd = MoveToAABone(obj, Frame, b_index)
    glTranslated 0, 0, obj.Bones(b_index).length
End Function
Function MoveToAABone(ByRef obj As AASkeleton, ByRef Frame As DAFrame, ByVal b_index As Integer) As Integer
    Dim BI As Integer

    Dim joint_stack() As Integer
    Dim jsp As Integer
    Dim rot_mat(16) As Double

    ReDim joint_stack(obj.NumBones * 4)
    jsp = 0

    joint_stack(jsp) = -1

    glMatrixMode GL_MODELVIEW

    glPushMatrix
    glTranslated Frame.X_start, Frame.Y_start, Frame.Z_start
    With Frame.Bones(0)
        BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
    End With
    glMultMatrixd rot_mat(0)
    For BI = 0 To b_index - 1
        glPushName BI
            While Not (obj.Bones(BI).ParentBone = joint_stack(jsp)) And jsp > 0
                glPopMatrix
                jsp = jsp - 1
            Wend
            glPushMatrix

            'glRotated Frame.Bones(bi + 1).Beta, 0#, 1#, 0#
            'glRotated Frame.Bones(bi + 1).Alpha, 1#, 0#, 0#
            'glRotated Frame.Bones(bi + 1).Gamma, 0#, 0#, 1#
            With Frame.Bones(BI + IIf(obj.NumBones > 1, 1, 0))
                BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
            End With
            glMultMatrixd rot_mat(0)

            glTranslated 0, 0, obj.Bones(BI).length

            jsp = jsp + 1
            joint_stack(jsp) = BI
        glPopName
    Next BI


    While Not (obj.Bones(BI).ParentBone = joint_stack(jsp)) And jsp > 0
        glPopMatrix
        jsp = jsp - 1
    Wend
    'glPopMatrix

    'With Frame.Bones(b_index + IIf(obj.NumBones > 1, 1, 0))
        'glRotated .Beta, 0#, 1#, 0#
        'glRotated .Alpha, 1#, 0#, 0#
        'glRotated .Gamma, 0#, 0#, 1#
        With Frame.Bones(b_index + IIf(obj.NumBones > 1, 1, 0))
            BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
        End With
        glMultMatrixd rot_mat(0)
    'End With

    MoveToAABone = jsp + 1
End Function
Sub ApplyAAChanges(ByRef obj As AASkeleton, ByRef Frame As DAFrame, ByRef FrameWeapon As DAFrame)
    Dim BI As Integer
    Dim wi As Integer
    Dim joint_stack() As Integer
    Dim jsp As Integer
    Dim rot_mat(16) As Double

    ReDim joint_stack(obj.NumBones)
    jsp = 0

    joint_stack(jsp) = -1

    glMatrixMode GL_MODELVIEW
    glPushMatrix
    'glLoadIdentity
    glTranslated Frame.X_start, Frame.Y_start, Frame.Z_start
    With Frame.Bones(0)
        BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
    End With
    glMultMatrixd rot_mat(0)
    For BI = 0 To obj.NumBones - 1
        While Not (obj.Bones(BI).ParentBone = joint_stack(jsp)) And jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend
        glPushMatrix

        'glRotated Frame.Bones(bi + 1).Beta, 0#, 1#, 0#
        'glRotated Frame.Bones(bi + 1).Alpha, 1#, 0#, 0#
        'glRotated Frame.Bones(bi + 1).Gamma, 0#, 0#, 1#
        With Frame.Bones(BI + IIf(obj.NumBones > 1, 1, 0))
            BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
        End With
        glMultMatrixd rot_mat(0)

        If obj.Bones(BI).hasModel Then _
            ApplyAABoneChanges obj.Bones(BI), 0

        glTranslated 0, 0, obj.Bones(BI).length

        jsp = jsp + 1
        joint_stack(jsp) = BI
    Next BI

    While jsp > 0
        glPopMatrix
        jsp = jsp - 1
    Wend
    glPopMatrix

    If obj.NumWeapons > 0 Then
        glMatrixMode GL_MODELVIEW
        glPushMatrix
        'glLoadIdentity
        glTranslated FrameWeapon.X_start, FrameWeapon.Y_start, FrameWeapon.Z_start
        glMultMatrixd rot_mat(0)
        With FrameWeapon.Bones(0)
            BuildRotationMatrixWithQuaternions .alpha, .Beta, .Gamma, rot_mat
        End With
        For wi = 0 To obj.NumWeapons - 1
            ApplyAAWeaponChanges obj.WeaponModels(wi), ComputeAADiameter(obj) * 4
        Next wi
        glPopMatrix
    End If
End Sub
Sub CreateCompatibleDAAnimationsPack(ByRef obj As AASkeleton, ByRef AnimationsPack As DAAnimationsPack)
    With AnimationsPack
        .NumAnimations = 1
        .NumBodyAnimations = 1
        .NumWeaponAnimations = 0
        ReDim .BodyAnimations(0)
        ReDim .WeaponAnimations(0)
        CreateCompatibleDAAnimationsPackAnimation obj, .BodyAnimations(0)
    End With
End Sub
Sub CreateCompatibleDAAnimationsPackAnimation(ByRef obj As AASkeleton, ByRef Anim As DAAnimation)
    With Anim
        .NumFrames1 = 1
        .NumFrames2 = 1
        ReDim .Frames(0)
        CreateCompatibleDAAnimationsPackAnimation1stFrame obj, .Frames(0)
    End With
End Sub
Sub CreateCompatibleDAAnimationsPackAnimation1stFrame(ByRef obj As AASkeleton, ByRef Frame As DAFrame)
    Dim BI As Integer
    Dim last_joint As String
    Dim joint_stack() As String
    Dim jsp As Integer
    Dim jsp0 As Integer


    Dim StageIndex As Integer

    Dim HipArmAngle As Single
    Dim c1 As Single
    Dim c2 As Single

    ReDim joint_stack(obj.NumBones)
    jsp = 0
    jsp0 = 0

    joint_stack(jsp) = -1

    ReDim Frame.Bones(obj.NumBones)

    StageIndex = 1

    For BI = 0 To obj.NumBones - 1
        While Not (obj.Bones(BI).ParentBone = joint_stack(jsp)) And jsp > 0
            glPopMatrix
            jsp = jsp - 1
        Wend

        If jsp0 > jsp Then
            StageIndex = StageIndex + 1
        End If
        ''Debug.Print obj.Bones(bi + 1).ParentBone, bi, jsp, StageIndex

        With Frame.Bones(BI + 1)
            Select Case StageIndex
                Case 1:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                    If BI = 1 Then StageIndex = 2
                Case 2:
                    .alpha = -145
                    .Beta = 0
                    .Gamma = 0
                    StageIndex = 3
                Case 3:
                    If jsp > jsp0 Then
                        .alpha = 0
                        .Beta = 0
                        .Gamma = 0
                    Else
                        .alpha = -180
                        .Beta = 0
                        .Gamma = 180
                        StageIndex = 5
                    End If
                Case 4:
                    .alpha = -180
                    .Beta = 0
                    .Gamma = 180
                    StageIndex = 5
                Case 5:
                    .alpha = 0
                    .Beta = 90
                    .Gamma = 0
                    StageIndex = 6
                Case 6:
                    .alpha = 0
                    .Beta = -60
                    .Gamma = 0
                    StageIndex = 7
                Case 7:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                    StageIndex = 8
                Case 8:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                    StageIndex = 9
                Case 9:
                    .alpha = -90
                    .Beta = 0
                    .Gamma = 0
                    StageIndex = 10
                Case 10:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                Case 11:
                    .alpha = 0
                    .Beta = -90
                    .Gamma = 0
                    StageIndex = 12
                Case 12:
                    .alpha = 0
                    .Beta = 60
                    .Gamma = 0
                    StageIndex = 13
                Case 13:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                    StageIndex = 14
                Case 14:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                    StageIndex = 15
                Case 15:
                    .alpha = -90
                    .Beta = 0
                    .Gamma = 0
                    StageIndex = 16
                Case 16:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                Case 17:
                    c1 = obj.Bones(BI + 1).length - obj.Bones(BI).length * 0.01
                    c2 = Sqr(obj.Bones(BI + 1).length ^ 2 - c1 ^ 2)
                    HipArmAngle = Atn(c2 / c1) / PI_180
                    .alpha = 0
                    .Beta = HipArmAngle
                    .Gamma = 0
                    StageIndex = 18
                Case 18:
                    .alpha = 0
                    .Beta = -HipArmAngle - 90
                    .Gamma = 0
                    StageIndex = 19
                Case 19:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                Case 20:
                    .alpha = 0
                    .Beta = -HipArmAngle
                    .Gamma = 0
                    StageIndex = 21
                Case 21:
                    .alpha = 0
                    .Beta = HipArmAngle + 90
                    .Gamma = 0
                    StageIndex = 22
                Case 22:
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
                Case 23:
                    .alpha = 90
                    .Beta = 0
                    .Gamma = 0
                Case Else
                    .alpha = 0
                    .Beta = 0
                    .Gamma = 0
            End Select
        End With
        jsp0 = jsp
        jsp = jsp + 1
        joint_stack(jsp) = BI

    Next BI
    Frame.Bones(1).Gamma = 180
    Frame.Bones(0).alpha = 90

End Sub

Public Function GetBattleModelTextureFilename(ByRef obj As AASkeleton, ByVal tex_num As Integer) As String
    GetBattleModelTextureFilename = LCase$(Left(obj.filename, 2)) + "a" + Chr$(99 + tex_num)
End Function
Sub InterpolateDAAnimationsPack(ByRef skeleton As AASkeleton, ByRef anims_pack As DAAnimationsPack, ByVal num_interpolated_frames As Integer, ByVal is_loopQ As Boolean)
    Dim ai As Integer

    With anims_pack
        For ai = 0 To .NumBodyAnimations - 1
            If .BodyAnimations(ai).NumFrames2 > 1 Then
                InterpolateBodyDAAnimation skeleton, .BodyAnimations(ai), num_interpolated_frames, is_loopQ

                If ai < .NumWeaponAnimations And skeleton.NumWeapons > 0 Then
                    InterpolateWeaponDAAnimation skeleton, .WeaponAnimations(ai), num_interpolated_frames, is_loopQ, .BodyAnimations(ai).NumFrames1, .BodyAnimations(ai).NumFrames2
                End If
            End If
        Next ai
    End With
End Sub
Sub InterpolateBodyDAAnimation(ByRef skeleton As AASkeleton, ByRef Anim As DAAnimation, ByVal num_interpolated_frames As Integer, ByVal is_loopQ As Boolean)
    Dim primary_secondary_counters_coef As Single
    Dim next_elem_diff As Integer
    Dim frame_offset As Integer
    Dim fi As Integer
    Dim ifi As Integer
    Dim base_final_frame As Integer
    Dim alpha As Single

    next_elem_diff = num_interpolated_frames + 1

    frame_offset = 0
    If Not is_loopQ Then
        frame_offset = num_interpolated_frames
    End If

    With Anim
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
                If skeleton.NumBones > 1 Then
                    GetTwoDAFramesInterpolation skeleton, .Frames(fi), .Frames(fi + num_interpolated_frames + 1), alpha, .Frames(fi + ifi)
                Else
                    GetTwoDAFramesWeaponInterpolation skeleton, .Frames(fi), .Frames(fi + num_interpolated_frames + 1), alpha, .Frames(fi + ifi)
                End If
            Next ifi
        Next fi

        base_final_frame = .NumFrames2 - num_interpolated_frames - 1
        If is_loopQ Then
            For ifi = 1 To num_interpolated_frames
                alpha = CSng(ifi) / CSng(num_interpolated_frames + 1)
                If skeleton.NumBones > 1 Then
                    GetTwoDAFramesInterpolation skeleton, .Frames(base_final_frame), .Frames(0), alpha, .Frames(base_final_frame + ifi)
                Else
                    GetTwoDAFramesWeaponInterpolation skeleton, .Frames(base_final_frame), .Frames(0), alpha, .Frames(base_final_frame + ifi)
                End If
            Next ifi
        End If

        'NormalizeDAAnimationsPackAnimation Anim
    End With
End Sub
Sub InterpolateWeaponDAAnimation(ByRef skeleton As AASkeleton, ByRef Anim As DAAnimation, ByVal num_interpolated_frames As Integer, ByVal is_loopQ As Boolean, ByVal body_num_frames1 As Integer, ByVal body_num_frames2 As Integer)
    Dim next_elem_diff As Integer
    Dim frame_offset As Integer
    Dim fi As Integer
    Dim ifi As Integer
    Dim base_final_frame As Integer
    Dim alpha As Single

    next_elem_diff = num_interpolated_frames + 1

    frame_offset = 0
    If Not is_loopQ Then
        frame_offset = num_interpolated_frames
    End If

    With Anim
        .NumFrames2 = body_num_frames2
        .NumFrames1 = body_num_frames1

        ReDim Preserve .Frames(.NumFrames2 - 1)
        'Move the original frames into their new positions
        For fi = .NumFrames2 - (1 + num_interpolated_frames - frame_offset) To 0 Step -next_elem_diff
            .Frames(fi) = .Frames(fi / (num_interpolated_frames + 1))
        Next fi

        'Interpolate the new frames
        For fi = 0 To .NumFrames2 - (1 + num_interpolated_frames + num_interpolated_frames - frame_offset) Step next_elem_diff
            For ifi = 1 To num_interpolated_frames
                alpha = CSng(ifi) / CSng(num_interpolated_frames + 1)
                GetTwoDAFramesWeaponInterpolation skeleton, .Frames(fi), .Frames(fi + next_elem_diff), alpha, .Frames(fi + ifi)
            Next ifi
        Next fi

        base_final_frame = .NumFrames2 - num_interpolated_frames - 1
        If is_loopQ Then
            For ifi = 1 To num_interpolated_frames
                alpha = CSng(ifi) / CSng(num_interpolated_frames + 1)
                GetTwoDAFramesWeaponInterpolation skeleton, .Frames(base_final_frame), .Frames(0), alpha, .Frames(base_final_frame + ifi)
            Next ifi
        End If
    End With
End Sub
Sub GetTwoDAFramesInterpolation(ByRef skeleton As AASkeleton, ByRef frame_a As DAFrame, ByRef frame_b As DAFrame, ByVal alpha As Single, ByRef frame_out As DAFrame)
    Dim BI As Integer
    Dim last_joint As Integer
    Dim joint_stack() As Integer
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
        num_bones = skeleton.NumBones 'frame_a.NumBones
        If num_bones <> UBound(frame_a.Bones) Then
            Debug.Assert "XXX"
        End If

        If num_bones = 1 Then
            GetTwoDAFramesWeaponInterpolation skeleton, frame_a, frame_b, alpha, frame_out
        Else
            ReDim .Bones(num_bones)
            '.NumBones = frame_a.NumBones

            alpha_inv = 1# - alpha
            .X_start = frame_a.X_start * alpha_inv + frame_b.X_start * alpha
            .Y_start = frame_a.Y_start * alpha_inv + frame_b.Y_start * alpha
            .Z_start = frame_a.Z_start * alpha_inv + frame_b.Z_start * alpha

            ReDim joint_stack(num_bones)
            ReDim rotations_stack_a(num_bones)
            ReDim rotations_stack_b(num_bones)
            ReDim rotations_stack_acum(num_bones)

            rotations_stack_a(0) = GetQuaternionFromEulerYXZr(frame_a.Bones(0).alpha, frame_a.Bones(0).Beta, frame_a.Bones(0).Gamma)
            NormalizeQuaternion rotations_stack_a(0)
            rotations_stack_b(0) = GetQuaternionFromEulerYXZr(frame_b.Bones(0).alpha, frame_b.Bones(0).Beta, frame_b.Bones(0).Gamma)
            NormalizeQuaternion rotations_stack_b(0)
            rotations_stack_acum(0) = QuaternionSlerp2(rotations_stack_a(0), rotations_stack_b(0), alpha)
            NormalizeQuaternion rotations_stack_acum(0)

            joint_stack(0) = -1
            jsp = 0
            For BI = 0 To num_bones - 1
                While jsp > 0 And skeleton.Bones(BI).ParentBone <> joint_stack(jsp)
                    jsp = jsp - 1
                Wend

                quat_a = GetQuaternionFromEulerYXZr(frame_a.Bones(BI + 1).alpha, frame_a.Bones(BI + 1).Beta, frame_a.Bones(BI + 1).Gamma)
                NormalizeQuaternion quat_a
                quat_b = GetQuaternionFromEulerYXZr(frame_b.Bones(BI + 1).alpha, frame_b.Bones(BI + 1).Beta, frame_b.Bones(BI + 1).Gamma)
                NormalizeQuaternion quat_b

                MultiplyQuaternions rotations_stack_a(jsp), quat_a, quat_acum_a
                NormalizeQuaternion quat_acum_a
                rotations_stack_a(jsp + 1) = quat_acum_a
                MultiplyQuaternions rotations_stack_b(jsp), quat_b, quat_acum_b
                NormalizeQuaternion quat_acum_b
                rotations_stack_b(jsp + 1) = quat_acum_b

                quat_interp = QuaternionSlerp2(quat_acum_a, quat_acum_b, alpha)
                rotations_stack_acum(jsp + 1) = quat_interp
                quat_acum_inverse = GetQuaternionConjugate(rotations_stack_acum(jsp))
                MultiplyQuaternions quat_acum_inverse, quat_interp, quat_interp_final
                NormalizeQuaternion quat_interp_final

                BuildMatrixFromQuaternion quat_interp_final, mat
                euler_res = GetEulerYXZrFromMatrix(mat)

                .Bones(BI + 1).alpha = euler_res.y
                .Bones(BI + 1).Beta = euler_res.x
                .Bones(BI + 1).Gamma = euler_res.z

                jsp = jsp + 1
                joint_stack(jsp) = BI
            Next BI

            BuildMatrixFromQuaternion rotations_stack_acum(0), mat
            euler_res = GetEulerYXZrFromMatrix(mat)

            .Bones(0).alpha = euler_res.y
            .Bones(0).Beta = euler_res.x
            .Bones(0).Gamma = euler_res.z
        End If
    End With
End Sub

Sub GetTwoDAFramesWeaponInterpolation(ByRef skeleton As AASkeleton, ByRef frame_a As DAFrame, ByRef frame_b As DAFrame, ByVal alpha As Single, ByRef frame_out As DAFrame)
    Dim quat_a As Quaternion
    Dim quat_b As Quaternion
    Dim quat_interp As Quaternion
    Dim euler_res As Point3D
    Dim alpha_inv As Single
    Dim mat(16) As Double

    With frame_out
        '.NumBones = frame_a.NumBones
        ReDim .Bones(0)

        alpha_inv = 1# - alpha
        .X_start = frame_a.X_start * alpha_inv + frame_b.X_start * alpha
        .Y_start = frame_a.Y_start * alpha_inv + frame_b.Y_start * alpha
        .Z_start = frame_a.Z_start * alpha_inv + frame_b.Z_start * alpha

        quat_a = GetQuaternionFromEulerYXZr(frame_a.Bones(0).alpha, frame_a.Bones(0).Beta, frame_a.Bones(0).Gamma)
        NormalizeQuaternion quat_a
        quat_b = GetQuaternionFromEulerYXZr(frame_b.Bones(0).alpha, frame_b.Bones(0).Beta, frame_b.Bones(0).Gamma)
        NormalizeQuaternion quat_b

        quat_interp = QuaternionSlerp2(quat_a, quat_b, alpha)
        NormalizeQuaternion quat_interp
        BuildMatrixFromQuaternion quat_interp, mat
        euler_res = GetEulerYXZrFromMatrix(mat)

        'NormalizeEulerAngles euler_res

        .Bones(0).alpha = euler_res.y
        .Bones(0).Beta = euler_res.x
        .Bones(0).Gamma = euler_res.z
    End With
End Sub
Function GetLimitCharacterFileName(ByVal limit_filename As String) As String
    Dim clean_filename As String
    Dim char_id As String

    clean_filename = LCase(TrimPath(limit_filename))
    char_id = Mid$(clean_filename, 4, 2)
    If char_id = "br" Or clean_filename = "hvshot.a00" Then
        GetLimitCharacterFileName = BARRET_BATTLE_SKELETON
    ElseIf char_id = "cd" Then
        GetLimitCharacterFileName = CID_BATTLE_SKELETON
    ElseIf char_id = "cl" Or clean_filename = "blaver.a00" Or clean_filename = "kyou.a00" Then
        GetLimitCharacterFileName = CLOUD_BATTLE_SKELETON
    ElseIf char_id = "ea" Or clean_filename = "iyash.a00" Or clean_filename = "kodo.a00" Then
        GetLimitCharacterFileName = AERITH_BATTLE_SKELETON
    ElseIf char_id = "rd" Or clean_filename = "limsled.a00" Then
        GetLimitCharacterFileName = RED_BATTLE_SKELETON
    ElseIf char_id = "yf" Then
        GetLimitCharacterFileName = YUFFIE_BATTLE_SKELETON
    ElseIf clean_filename = "limfast.a00" Then
        GetLimitCharacterFileName = TIFA_BATTLE_SKELETON
    ElseIf clean_filename = "dice.a00" Then
        GetLimitCharacterFileName = CAITSITH_BATTLE_SKELETON
    Else
        GetLimitCharacterFileName = ""
    End If
End Function

Function ModelHasLimitBreaks(ByVal model_filename As String) As Boolean
    Dim clean_filename As String
    Dim char_id As String

    clean_filename = LCase(TrimPath(model_filename))
    char_id = Mid$(clean_filename, 4, 2)

    'Barret
    If clean_filename = "sbaa" Or clean_filename = "scaa" Or clean_filename = "sdaa" Or clean_filename = "seaa" Then
        ModelHasLimitBreaks = True
    'Cloud
    ElseIf clean_filename = "siaa" Or clean_filename = "rtaa" Then
        ModelHasLimitBreaks = True
    'Cid
    ElseIf clean_filename = "rzaa" Then
        ModelHasLimitBreaks = True
    'Cait Sith
    ElseIf clean_filename = "ryaa" Then
        ModelHasLimitBreaks = True
    'Yuffie
    ElseIf clean_filename = "rxaa" Then
        ModelHasLimitBreaks = True
    'Red XIII
    ElseIf clean_filename = "rwaa" Then
        ModelHasLimitBreaks = True
    'Aerith
    ElseIf clean_filename = "rvaa" Then
        ModelHasLimitBreaks = True
    'Tifa
    ElseIf clean_filename = "ruaa" Then
        ModelHasLimitBreaks = True
    Else
        ModelHasLimitBreaks = False
    End If

End Function

Function GetModelAnimationPacksFilter(ByVal clean_filename As String) As String
    Dim char_id As String
    char_id = Mid$(clean_filename, 1, 2)
    'Barret
    If clean_filename = "sbaa" Or clean_filename = "scaa" Or clean_filename = "sdaa" Or clean_filename = "seaa" Then
        GetModelAnimationPacksFilter = "Limit breaks (limbr*.a00, hvshot.a00)|limbr2.a00;limbr3.a00;limbr4.a00;limbr5.a00;limbr6.a00;limbr7.a00;hvshot.a00"
    'Cloud
    ElseIf clean_filename = "siaa" Or clean_filename = "rtaa" Then
        GetModelAnimationPacksFilter = "Limit breaks (limcl*.a00, blaver.a00, kyou.a00)|limcl2.a00;limcl3.a00;limcl4.a00;limcl5.a00;limcl6.a00;limcl7.a00;blaver.a00;kyou.a00"
    'Cid
    ElseIf clean_filename = "rzaa" Then
        GetModelAnimationPacksFilter = "Limit breaks (limcd*.a00)|limcd1.a00;limcd2.a00;limcd3.a00;limcd4.a00;limcd5.a00;limcd6.a00"
    'Cait Sith
    ElseIf clean_filename = "ryaa" Then
        GetModelAnimationPacksFilter = "Limit breaks (dice.a00)|dice.a00"
    'Yuffie
    ElseIf clean_filename = "rxaa" Then
        GetModelAnimationPacksFilter = "Limit breaks (limyf*.a00)|limyf1.a00;limyf2.a00;limyf3.a00;limyf4.a00;limyf5.a00;limyf6.a00;limyf7.a00"
    'Red XIII
    ElseIf clean_filename = "rwaa" Then
        GetModelAnimationPacksFilter = "Limit breaks (limrd*.a00, limsled.a00)|limrd3.a00;limrd4.a00;limrd5.a00;limrd6.a00;limrd7.a00;limsled.a00"
    'Aerith
    ElseIf clean_filename = "rvaa" Then
        GetModelAnimationPacksFilter = "Limit breaks (limea*.a00, iyash.a00, kodo.a00)|limea2.a00;limea3.a00;limea4.a00;limea5.a00;limea6.a00;limea7.a00;iyash.a00;kodo.a00"
    'Tifa
    ElseIf clean_filename = "ruaa" Then
        GetModelAnimationPacksFilter = "Limit breaks (limfast.a00)|limfast.a00"
    Else
        GetModelAnimationPacksFilter = "ERROR!!!"
    End If

    GetModelAnimationPacksFilter = GetModelAnimationPacksFilter + "|Battle animations (" + char_id + "da)|" + char_id + "da"
End Function
