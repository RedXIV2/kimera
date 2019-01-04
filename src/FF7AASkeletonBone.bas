Attribute VB_Name = "FF7AASkeletonBone"
Option Explicit
Type AABone
    ParentBone As Long
    length As Single
    hasModel As Long
    Models() As PModel
'-------------Extra Atributes----------------
    NumModels As Integer
    ResizeX As Single
    ResizeY As Single
    ResizeZ As Single
End Type
Sub ReadAABone(ByVal NFile As Integer, ByVal offset As Long, ByVal modelName As String, ByRef bone As AABone, ByVal load_geometryQ As Boolean)
    With bone
        Get NFile, offset, .ParentBone
        Get NFile, offset + 4, .length
        Get NFile, offset + 4 * 2, .hasModel
        If Not (.hasModel = 0) Then
            If load_geometryQ Then
                ReDim .Models(0)
                .NumModels = 1
                ReadPModel .Models(0), modelName
            End If
        End If
        
        .ResizeX = 1
        .ResizeY = 1
        .ResizeZ = 1
    End With
End Sub
Sub WriteAABone(ByVal NFile As Integer, ByVal offset As Long, ByVal modelName As String, ByRef bone As AABone)
    With bone
        Put NFile, offset, .ParentBone
        Put NFile, offset + 4, .length
        Put NFile, offset + 4 * 2, .hasModel
        If Not (.hasModel = 0) Then _
                WritePModel .Models(0), modelName
        
        .ResizeX = 1
        .ResizeY = 1
        .ResizeZ = 1
    End With
End Sub
Sub ReadAABattleLocationPiece(ByRef Piece As AABone, ByVal bone_index As Integer, ByVal modelName As String)
    With Piece
        .ParentBone = bone_index
        .hasModel = 1
        ReDim .Models(0)
        .NumModels = 1
        ReadPModel .Models(0), modelName
        .length = ComputeDiameter(.Models(0).BoundingBox) / 2
        
        .ResizeX = 1
        .ResizeY = 1
        .ResizeZ = 1
    End With
End Sub
Sub CreateDListsFromAASkeletonBone(ByRef obj As AABone)
    Dim mi As Integer
    
    For mi = 0 To obj.NumModels - 1
        CreateDListsFromPModel obj.Models(mi)
    Next mi
End Sub
Sub FreeAABoneResources(ByRef obj As AABone)
    Dim mi As Integer
    
    If obj.hasModel Then
        For mi = 0 To obj.NumModels - 1
            FreePModelResources obj.Models(mi)
        Next mi
    End If
End Sub
Sub DrawAASkeletonBone(ByRef bone As AABone, ByRef tex_ids() As Long, ByVal UseDLists As Boolean)
    Dim mi As Integer
    
    glMatrixMode GL_MODELVIEW
    glPushMatrix
    With bone
        glScalef .ResizeX, .ResizeY, .ResizeZ
    End With
    
    If bone.hasModel > 0 Then
        If Not UseDLists Then
            For mi = 0 To bone.NumModels - 1
                glPushMatrix
                With bone.Models(mi)
                    glTranslatef .RepositionX, .RepositionY, .RepositionZ
                    
                    glRotated .RotateAlpha, 1#, 0#, 0#
                    glRotated .RotateBeta, 0#, 1#, 0#
                    glRotated .RotateGamma, 0#, 0#, 1#
                    
                    glScalef .ResizeX, .ResizeY, .ResizeZ
                End With
                DrawPModel bone.Models(mi), tex_ids, False
                glPopMatrix
            Next mi
        Else
            For mi = 0 To bone.NumModels - 1
                glPushMatrix
                With bone.Models(mi)
                    glTranslatef .RepositionX, .RepositionY, .RepositionZ
                    
                    glRotated .RotateAlpha, 1#, 0#, 0#
                    glRotated .RotateBeta, 0#, 1#, 0#
                    glRotated .RotateGamma, 0#, 0#, 1#
                    
                    glScalef .ResizeX, .ResizeY, .ResizeZ
                End With
                DrawPModelDLists bone.Models(mi), tex_ids
                glPopMatrix
            Next mi
        End If
    End If
    glPopMatrix
End Sub
Sub DrawAABoneBoundingBox(ByRef bone As AABone)
    Dim rot_mat(16) As Double
    
    glDisable GL_DEPTH_TEST
    glMatrixMode GL_MODELVIEW
    With bone
        glScalef .ResizeX, .ResizeY, .ResizeZ
    End With
    If bone.hasModel Then
        'With bone.Models(0)
        '    glTranslatef .RepositionX, .RepositionY, .RepositionZ
        '
        '    BuildMatrixFromQuaternion .RotationQuaternion, rot_mat
       '
        '    glMultMatrixd rot_mat(0)
        '
        '    glScalef .ResizeX, .ResizeY, .ResizeZ
        'End With
        With bone.Models(0).BoundingBox
            DrawBox .max_x, .max_y, .max_z, .min_x, .min_y, .min_z, 1, 0, 0
        End With
    Else
        glColor3f 0, 1, 0
        glBegin GL_LINES
            glVertex3f 0, 0, 0
            glVertex3f 0, 0, bone.length
        glEnd
    End If
    glEnable GL_DEPTH_TEST
End Sub
Sub DrawAAModelBoundingBox(ByRef bone As AABone)
    Dim mi As Integer
    
    Dim max_x As Single
    Dim max_y As Single
    Dim max_z As Single
    
    Dim min_x As Single
    Dim min_y As Single
    Dim min_z As Single
    
    glMatrixMode GL_MODELVIEW
    With bone
        glScalef .ResizeX, .ResizeY, .ResizeZ
    End With
    
    If bone.NumModels = 0 Then
        glDisable GL_DEPTH_TEST
        glColor3f 1, 0, 0
        glBegin GL_LINES
            glVertex3f 0, 0, 0
            glVertex3f 0, 0, -bone.length
        glEnd
        glEnable GL_DEPTH_TEST
    Else
        max_x = -INFINITY_SINGLE
        max_y = -INFINITY_SINGLE
        max_z = -INFINITY_SINGLE
        
        min_x = INFINITY_SINGLE
        min_y = INFINITY_SINGLE
        min_z = INFINITY_SINGLE
        
        For mi = 0 To bone.NumModels - 1
            With bone.Models(mi).BoundingBox
                If max_x < .max_x Then max_x = .max_x
                If max_y < .max_y Then max_y = .max_y
                If max_z < .max_z Then max_z = .max_z
                
                If min_x > .min_x Then min_x = .min_x
                If min_y > .min_y Then min_y = .min_y
                If min_z > .min_z Then min_z = .min_z
            End With
        Next mi
        
        glDisable GL_DEPTH_TEST
        DrawBox max_x, max_y, max_z, min_x, min_y, min_z, 1, 0, 0
        glEnable GL_DEPTH_TEST
    End If
End Sub
Sub DrawAABoneModelBoundingBox(ByRef bone As AABone, ByVal p_index As Integer)
    Dim rot_mat(16) As Double
    glDisable GL_DEPTH_TEST
    glMatrixMode GL_MODELVIEW
    With bone
        glScalef .ResizeX, .ResizeY, .ResizeZ
    End With
    
    With bone.Models(p_index)
        glTranslatef .RepositionX, .RepositionY, .RepositionZ
        
        BuildMatrixFromQuaternion .RotationQuaternion, rot_mat
        
        glMultMatrixd rot_mat(0)
        
        glScalef .ResizeX, .ResizeY, .ResizeZ
    End With
    With bone.Models(p_index).BoundingBox
        DrawBox .max_x, .max_y, .max_z, .min_x, .min_y, .min_z, 0, 1, 0
    End With
    glEnable GL_DEPTH_TEST
End Sub
Sub ApplyAABoneChanges(ByRef bone As AABone, ByVal diameter As Single)
    Dim mi As Integer

    For mi = 0 To bone.NumModels - 1
        If bone.hasModel Then
            If glIsEnabled(GL_LIGHTING) = GL_TRUE Then _
                ApplyCurrentVColors bone.Models(mi)
        
            glMatrixMode GL_MODELVIEW
            glPushMatrix
            With bone.Models(mi)
                SetCameraModelViewQuat .RepositionX, .RepositionY, _
                            .RepositionZ, .RotationQuaternion, _
                            .ResizeX, .ResizeY, .ResizeZ
            End With
                            
            glScalef bone.ResizeX, bone.ResizeY, bone.ResizeZ
            
            ApplyPChanges bone.Models(mi), True
            glMatrixMode GL_MODELVIEW
            glPopMatrix
        End If
    Next mi
    
    MergeAABoneModels bone
    If bone.NumModels > 1 Then
        ReDim Preserve bone.Models(0)
        bone.NumModels = 1
    End If
End Sub
Sub ApplyAAWeaponChanges(ByRef weapon As PModel, ByVal diameter As Single)
    'If glIsEnabled(GL_LIGHTING) = GL_TRUE Then _
    '    ApplyCurrentVColors weapon

    'glMatrixMode GL_MODELVIEW
    'glPushMatrix
    'glMatrixMode GL_PROJECTION
    'glPushMatrix
    'With weapon
    '    SetCameraPModel weapon, .RepositionX, .RepositionY, _
    '                    .RepositionZ + ComputeDiameter(.BoundingBox) * 2, _
    '                    .RotateAlpha, .RotateBeta, .RotateGamma, _
    '                    .ResizeX, .ResizeY, .ResizeZ
    'End With
   '
   ' ApplyPChanges weapon, True
   ' glMatrixMode GL_PROJECTION
   ' glPopMatrix
   ' glMatrixMode GL_MODELVIEW
   ' glPopMatrix
   Dim mi As Integer

    If glIsEnabled(GL_LIGHTING) = GL_TRUE Then _
        ApplyCurrentVColors weapon

    glMatrixMode GL_MODELVIEW
    glPushMatrix
    With weapon
        SetCameraModelView .RepositionX, .RepositionY, _
                    .RepositionZ, .RotateAlpha, .RotateBeta, .RotateGamma, _
                    .ResizeX, .ResizeY, .ResizeZ
    End With
                    
    glScalef weapon.ResizeX, weapon.ResizeY, weapon.ResizeZ
    
    ApplyPChanges weapon, True
    glMatrixMode GL_MODELVIEW
    glPopMatrix
End Sub
Sub MergeAABoneModels(ByRef bone As AABone)
    Dim mi As Integer
    
    With bone
        For mi = 1 To .NumModels - 1
            MergePModels .Models(0), .Models(mi)
        Next mi
    End With
End Sub
Sub AddAABoneModel(ByRef bone As AABone, ByRef Piece As PModel)
    With bone
        .NumModels = .NumModels + 1
        ReDim Preserve .Models(.NumModels)
        .Models(.NumModels - 1) = Piece
        If .NumModels > 1 Then _
            .Models(.NumModels - 1).filename = Left$(.Models(0).filename, Len(.Models(0).filename) - 2) + Right$(Str$(.NumModels - 1), Len(Str$(.NumModels - 1)) - 1) + ".P"
        '.Models(.NumModels - 1). = .Resources(0).res_file + Right$(Str$(.NumResources - 1), Len(Str$(.NumResources - 1)) - 1)
        .hasModel = 1
    End With
End Sub
Sub RemoveAABoneModel(ByRef bone As AABone, ByVal m_index As Integer)
    Dim mi As Integer
    
    With bone
        If m_index < .NumModels - 1 Then
            For mi = m_index To .NumModels - 2
                .Models(mi) = .Models(mi + 1)
            Next mi
        End If
        .NumModels = .NumModels - 1
        ReDim Preserve .Models(.NumModels)
        If .NumModels <= 0 Then .hasModel = 0
    End With
End Sub
Sub ComputeAABoneBoundingBox(ByRef bone As AABone, ByRef p_min As Point3D, ByRef p_max As Point3D)
    Dim mi As Integer
    Dim p_min_aux As Point3D
    Dim p_max_aux As Point3D
    Dim p_min_aux_trans As Point3D
    Dim p_max_aux_trans As Point3D
    Dim MV_matrix(16) As Double
    
    glMatrixMode GL_MODELVIEW
    glPushMatrix
    glLoadIdentity
    With bone
        glScalef .ResizeX, .ResizeY, .ResizeZ
    End With
    If bone.hasModel Then
        p_max.x = -INFINITY_SINGLE
        p_max.y = -INFINITY_SINGLE
        p_max.z = -INFINITY_SINGLE
        
        p_min.x = INFINITY_SINGLE
        p_min.y = INFINITY_SINGLE
        p_min.z = INFINITY_SINGLE
        For mi = 0 To bone.NumModels - 1
            glPushMatrix
            With bone.Models(mi)
                glTranslatef .RepositionX, .RepositionY, .RepositionZ
                
                glRotated .RotateBeta, 0#, 1#, 0#
                glRotated .RotateAlpha, 1#, 0#, 0#
                glRotated .RotateGamma, 0#, 0#, 1#
                
                glScalef .ResizeX, .ResizeY, .ResizeZ
            End With
            With bone.Models(mi).BoundingBox
                p_min_aux.x = .min_x
                p_min_aux.y = .min_y
                p_min_aux.z = .min_z
                
                p_max_aux.x = .max_x
                p_max_aux.y = .max_y
                p_max_aux.z = .max_z
            End With
            
            glGetDoublev GL_MODELVIEW_MATRIX, MV_matrix(0)
            
            ComputeTransformedBoxBoundingBox MV_matrix, p_min_aux, p_max_aux, _
                p_min_aux_trans, p_max_aux_trans
            
            With p_max_aux_trans
                If p_max.x < .x Then p_max.x = .x
                If p_max.y < .y Then p_max.y = .y
                If p_max.z < .z Then p_max.z = .z
            End With
                
            With p_min_aux_trans
                If p_min.x > .x Then p_min.x = .x
                If p_min.y > .y Then p_min.y = .y
                If p_min.z > .z Then p_min.z = .z
            End With
            glPopMatrix
        Next mi
    Else
        p_max.x = 0
        p_max.y = 0
        p_max.z = 0
        
        p_min.x = 0
        p_min.y = 0
        p_min.z = 0
    End If
    glPopMatrix
End Sub

