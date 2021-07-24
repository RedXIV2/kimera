Attribute VB_Name = "FF7HRCSkeletonBone"
Option Explicit
Type HRCBone
    NumResources As Integer
    Resources() As RSBResource
    length As Double
    joint_i As String
    joint_f As String
'-------------Extra Atributes----------------
    ResizeX As Single
    ResizeY As Single
    ResizeZ As Single
End Type
Sub ReadHRCBone(ByVal NFile As Integer, ByRef bone As HRCBone, ByRef textures_pool() As TEXTexture, ByVal load_geometryQ As Boolean)
    Dim ci As Integer
    Dim ri As Integer
    Dim line As String
    Dim file As String
    Dim NFileAux As Integer
    Dim char As String

    NFileAux = FreeFile

    With bone
        'Skip all empty lines or comments

        'Bone name
        Do
             Line Input #NFile, .joint_i
        Loop While Len(.joint_i) = 0 Or Left$(.joint_i, 1) = "#"

        'Root name
        Do
            Line Input #NFile, .joint_f
        Loop While Len(.joint_f) = 0 Or Left$(.joint_f, 1) = "#"

        'Bone length
        Do
            Line Input #NFile, line
        Loop While Len(line) = 0 Or Left$(line, 1) = "#"
        .length = val(line)

        'RSB list
        Do
            Line Input #NFile, line
        Loop While Len(line) = 0 Or Left$(line, 1) = "#"

        'Parse RSB list
        If load_geometryQ Then
            If val(Left$(line, 1)) > 0 Then
                ci = 1
                While (IsNumeric(Left$(line, ci)))
                    .NumResources = val(Left$(line, ci))
                    ci = ci + 1
                Wend
                ReDim .Resources(.NumResources)

                For ci = ci To Len(line)
                    char = Mid$(line, ci, 1)

                    If char <> " " Then
                        file = file + char
                    Else
                        ReadRSBResource bone.Resources(ri), file, textures_pool
                        file = ""
                        ri = ri + 1
                    End If
                Next ci

                If file <> "" Then ReadRSBResource bone.Resources(ri), file, textures_pool

                .ResizeX = 1
                .ResizeY = 1
                .ResizeZ = 1
            End If
        End If
    End With
End Sub
Sub WriteHRCBone(ByVal NFile As Integer, ByRef bone As HRCBone)
    Dim ri As Integer
    Dim rsd_list As String
    Dim base_name_rsd As String
    Dim base_name_p As String

    With bone
        Print #NFile, ""
        Print #NFile, .joint_i
        Print #NFile, .joint_f
        Print #NFile, Right$(Str$(.length), Len(Str$(.length)) - 1)

        rsd_list = .NumResources

        If (.NumResources > 0) Then
            base_name_rsd = Left$(.Resources(0).res_file, 4)
            base_name_p = Left$(.Resources(0).Model.filename, 4)

            rsd_list = rsd_list + " " + .Resources(0).res_file
            WriteRSBResource .Resources(0), .Resources(0).res_file
            WritePModel .Resources(0).Model, .Resources(0).Model.filename
            For ri = 1 To .NumResources - 1
                .Resources(ri).res_file = base_name_rsd + Right$(Str$(ri), Len(Str$(ri)) - 1)
                .Resources(ri).Model.filename = base_name_p + Right$(Str$(ri), Len(Str$(ri)) - 1) + ".p"
                rsd_list = rsd_list + " " + .Resources(ri).res_file
                WriteRSBResource .Resources(ri), .Resources(ri).res_file
                WritePModel .Resources(ri).Model, .Resources(ri).Model.filename
            Next ri

            If Len(rsd_list) = 1 Then rsd_list = rsd_list + " "
        End If

        Print #NFile, rsd_list

    End With
End Sub
Sub CreateDListsFromHRCSkeletonBone(ByRef obj As HRCBone)
    Dim ri As Integer

    With obj
        For ri = 0 To .NumResources - 1
            CreateDListsFromRSBResource .Resources(ri)
        Next ri
    End With
End Sub
Sub FreeHRCBoneResources(ByRef obj As HRCBone)
    Dim ri As Integer

    With obj
        For ri = 0 To .NumResources - 1
            FreeRSBResourceResources .Resources(ri)
        Next ri
    End With
End Sub
Sub DrawHRCBone(ByRef bone As HRCBone, ByVal UseDLists As Boolean)
    Dim ri As Integer

    glMatrixMode GL_MODELVIEW
    glPushMatrix
    With bone
        glScalef .ResizeX, .ResizeY, .ResizeZ
    End With

    For ri = 0 To bone.NumResources - 1
        DrawRSBResource bone.Resources(ri), UseDLists
    Next ri
    glPopMatrix
End Sub
Sub DrawHRCBoneBoundingBox(ByRef bone As HRCBone)
    Dim ri As Integer

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

    If bone.NumResources = 0 Then
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

        For ri = 0 To bone.NumResources - 1
            With bone.Resources(ri).Model.BoundingBox
                If max_x < .max_x Then max_x = .max_x
                If max_y < .max_y Then max_y = .max_y
                If max_z < .max_z Then max_z = .max_z

                If min_x > .min_x Then min_x = .min_x
                If min_y > .min_y Then min_y = .min_y
                If min_z > .min_z Then min_z = .min_z
            End With
        Next ri

        glDisable GL_DEPTH_TEST
        DrawBox max_x, max_y, max_z, min_x, min_y, min_z, 1, 0, 0
        glEnable GL_DEPTH_TEST
    End If
End Sub
Sub DrawHRCBonePieceBoundingBox(ByRef bone As HRCBone, ByVal p_index As Integer)
    Dim rot_mat(16) As Double

    glDisable GL_DEPTH_TEST
    glMatrixMode GL_MODELVIEW
    With bone
        glScalef .ResizeX, .ResizeY, .ResizeZ
    End With

    With bone.Resources(p_index).Model
        glTranslatef .RepositionX, .RepositionY, .RepositionZ

        BuildMatrixFromQuaternion .RotationQuaternion, rot_mat

        glMultMatrixd rot_mat(0)

        glScalef .ResizeX, .ResizeY, .ResizeZ
    End With
    With bone.Resources(p_index).Model.BoundingBox
        DrawBox .max_x, .max_y, .max_z, .min_x, .min_y, .min_z, 0, 1, 0
    End With
    glEnable GL_DEPTH_TEST
End Sub
Sub AddHRCBonePiece(ByRef bone As HRCBone, ByRef Piece As PModel)
    With bone
        .NumResources = .NumResources + 1
        ReDim Preserve .Resources(.NumResources - 1)
        .Resources(.NumResources - 1).id = .Resources(0).id
        .Resources(.NumResources - 1).Model = Piece
        If .NumResources > 1 Then _
            .Resources(.NumResources - 1).Model.filename = Left$(.Resources(0).Model.filename, Len(.Resources(0).Model.filename) - 2) + Right$(Str$(.NumResources - 1), Len(Str$(.NumResources - 1)) - 1) + ".P"
        .Resources(.NumResources - 1).res_file = .Resources(0).res_file + Right$(Str$(.NumResources - 1), Len(Str$(.NumResources - 1)) - 1)
    End With
End Sub
Sub RemoveHRCBonePiece(ByRef bone As HRCBone, ByVal p_index As Integer)
    Dim pi As Integer

    With bone
        If p_index < .NumResources - 1 Then
            For pi = p_index To .NumResources - 2
                .Resources(pi) = .Resources(pi + 1)
            Next pi
        End If
        .NumResources = .NumResources - 1
        ReDim Preserve .Resources(.NumResources)
    End With
End Sub
Sub ApplyHRCBoneChanges(ByRef bone As HRCBone, ByVal diameter As Single, ByVal merge As Boolean)
    Dim ri As Integer

    For ri = 0 To bone.NumResources - 1
        ''Debug.Print "File=", bone.Resources(ri).res_file, bone.Resources(ri).Model.fileName
        If glIsEnabled(GL_LIGHTING) = GL_TRUE Then
            ApplyCurrentVColors bone.Resources(ri).Model
        End If

        glMatrixMode GL_MODELVIEW
        glPushMatrix
        With bone.Resources(ri).Model
            SetCameraModelViewQuat .RepositionX, .RepositionY, _
                            .RepositionZ, .RotationQuaternion, _
                            .ResizeX, .ResizeY, .ResizeZ
        End With

        glScalef bone.ResizeX, bone.ResizeY, bone.ResizeZ

        ApplyPChanges bone.Resources(ri).Model, False
        glMatrixMode GL_MODELVIEW
        glPopMatrix
    Next ri

    If merge Then
        MergeResources bone
        If bone.NumResources > 1 Then
            ReDim Preserve bone.Resources(0)
            bone.NumResources = 1
            ComputeBoundingBox bone.Resources(0).Model
        End If
    End If
End Sub
Sub MergeResources(ByRef bone As HRCBone)
    Dim ri As Integer

    For ri = 1 To bone.NumResources - 1
        MergeRSBResources bone.Resources(0), bone.Resources(ri)
    Next ri
End Sub
Function ComputeHRCBoneDiameter(ByRef bone As HRCBone) As Single
    Dim ri As Integer

    Dim p_max As Point3D
    Dim p_min As Point3D

    If bone.NumResources = 0 Then
        ComputeHRCBoneDiameter = 0
    Else
        p_max.x = -INFINITY_SINGLE
        p_max.y = -INFINITY_SINGLE
        p_max.z = -INFINITY_SINGLE

        p_min.x = INFINITY_SINGLE
        p_min.y = INFINITY_SINGLE
        p_min.z = INFINITY_SINGLE

        For ri = 0 To bone.NumResources - 1
            With bone.Resources(ri).Model.BoundingBox
                If p_max.x < .max_x Then p_max.x = .max_x
                If p_max.y < .max_y Then p_max.y = .max_y
                If p_max.z < .max_z Then p_max.z = .max_z

                If p_min.x > .min_x Then p_min.x = .min_x
                If p_min.y > .min_y Then p_min.y = .min_y
                If p_min.z > .min_z Then p_min.z = .min_z
            End With
        Next ri
        ComputeHRCBoneDiameter = CalculateDistance(p_max, p_min)
    End If
End Function

Sub ComputeHRCBoneBoundingBox(ByRef bone As HRCBone, ByRef p_min As Point3D, ByRef p_max As Point3D)
    Dim ri As Integer
    Dim p_min_part As Point3D
    Dim p_max_part As Point3D

    If bone.NumResources = 0 Then
        p_max.x = 0
        p_max.y = 0
        p_max.z = 0

        p_min.x = 0
        p_min.y = 0
        p_min.z = 0
    Else
        p_max.x = -INFINITY_SINGLE
        p_max.y = -INFINITY_SINGLE
        p_max.z = -INFINITY_SINGLE

        p_min.x = INFINITY_SINGLE
        p_min.y = INFINITY_SINGLE
        p_min.z = INFINITY_SINGLE

        For ri = 0 To bone.NumResources - 1
            ComputePModelBoundingBox bone.Resources(ri).Model, p_min_part, p_max_part
            With p_max_part
                If p_max.x < .x Then p_max.x = .x
                If p_max.y < .y Then p_max.y = .y
                If p_max.z < .z Then p_max.z = .z
            End With
            With p_min_part
                If p_min.x > .x Then p_min.x = .x
                If p_min.y > .y Then p_min.y = .y
                If p_min.z > .z Then p_min.z = .z
            End With
        Next ri
    End If
End Sub
