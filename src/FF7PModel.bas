Attribute VB_Name = "FF7PModel"
Option Explicit
Type PModel
    fileName As String
    head As PHeader
    Verts() As Point3D
    Normals() As Point3D
    TexCoords() As Point2D
    vcolors() As color
    PColors() As color
    Edges() As PEdge
    polys() As PPolygon
    hundrets() As PHundret
    Groups() As PGroup
    BoundingBox As PBoundingBox
    NormalIndex() As Long
'-------------Extra Atributes----------------
    ResizeX As Single
    ResizeY As Single
    ResizeZ As Single
    RotateAlpha As Single
    RotateBeta As Single
    RotateGamma As Single
    RotationQuaternion As Quaternion
    RepositionX As Single
    RepositionY As Single
    RepositionZ As Single
    diameter As Single
    DListNum As Long
End Type
'----------------------------------------------------------------------------------------------------
'=============================================BASIC I/O==============================================
'----------------------------------------------------------------------------------------------------
Sub ReadPModel(ByRef obj As PModel, ByVal fileName As String)
    Dim fileNumber As Integer
    
    On Error GoTo ErrHandRead
    
    If FileExist(fileName) Then
        fileNumber = FreeFile
        Open fileName For Binary As fileNumber
        ''Debug.Print fileName
        
        With obj
            .fileName = TrimPath(fileName)
            ReadHeader fileNumber, .head
            If .head.NumVerts <= 0 Then
                MsgBox fileName + ":Not a valid P file!!!", vbOKOnly, "Error reading"
            End If
            ReadVerts fileNumber, .Verts, .head.NumVerts
            ReadNormals fileNumber, &H81 + .head.NumVerts * 12, .Normals, .head.NumNormals
            ReadTexCoords fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12, .TexCoords, .head.NumTexCs
            ReadVColors fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8, .vcolors, .head.NumVerts
            ReadPColors fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + .head.NumVerts * 4, .PColors, .head.NumPolys
            ReadEdges fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys) * 4, .Edges, .head.NumEdges
            ReadPolygons fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys + .head.NumEdges) * 4, .polys, .head.NumPolys
            ReadHundrets fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys + .head.NumEdges) * 4 + .head.NumPolys * 24, .hundrets, .head.mirex_h
            ReadGroups fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys + .head.NumEdges) * 4 + .head.NumPolys * 24 + .head.mirex_h * 100, .Groups, .head.NumGroups
            ReadBoundingBox fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys + .head.NumEdges) * 4 + .head.NumPolys * 24 + .head.mirex_h * 100 + .head.NumGroups * 56 + 4, .BoundingBox
            ReadNormalIndex fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys + .head.NumEdges) * 4 + .head.NumPolys * 24 + .head.mirex_h * 100 + .head.NumGroups * 56 + 4 + 24, .NormalIndex, .head.NumNormInds
            
            .ResizeX = 1
            .ResizeY = 1
            .ResizeZ = 1
            .RotateAlpha = 0
            .RotateBeta = 0
            .RotateGamma = 0
            .RotationQuaternion.x = 0
            .RotationQuaternion.y = 0
            .RotationQuaternion.z = 0
            .RotationQuaternion.w = 1
            .RepositionX = 0
            .RepositionY = 0
            .RepositionZ = 0
            .diameter = ComputeDiameter(.BoundingBox)
        End With
        
        Close fileNumber
        'KillCrappyPolygons obj
        CheckModelConsistency obj
        KillUnusedVertices obj
        ComputeBoundingBox obj
        ComputeNormals obj
        CreateDListsFromPModel obj
    Else
        'Debug.Print "P file not found!!!"
        MsgBox "P file " + fileName + " not found!!!", vbOKOnly, "Error reading"
    End If
    Exit Sub
ErrHandRead:
    'Debug.Print "Error reading P file!!!"
    MsgBox "Error reading P file " + fileName + "!!!", vbOKOnly, "Error reading"
    ZeroMemory obj, Len(obj)
End Sub
Sub WritePModel(ByRef obj As PModel, ByVal fileName As String)
    Dim fileNumber As Integer
    
    On Error GoTo ErrHandWrite
    
    fileNumber = FreeFile
    Open fileName For Output As fileNumber
    Close fileNumber
    Open fileName For Binary As fileNumber
    
    With obj
        'If (LCase(Right(fileName, 2)) = ".p") Then
            '.head.VertexColor = 1
        'End If
        SetVColorsAlphaMAX .vcolors
        
        WriteHeader fileNumber, .head
        WriteVerts fileNumber, .Verts
        WriteNormals fileNumber, &H81 + .head.NumVerts * 12, .Normals
        WriteTexCoords fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12, .TexCoords
        WriteVColors fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8, .vcolors
        WritePColors fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + .head.NumVerts * 4, .PColors
        WriteEdges fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys) * 4, .Edges
        WritePolygons fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys + .head.NumEdges) * 4, .polys
        WriteHundrets fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys + .head.NumEdges) * 4 + .head.NumPolys * 24, .hundrets
        WriteGroups fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys + .head.NumEdges) * 4 + .head.NumPolys * 24 + .head.mirex_h * 100, .Groups
        WriteBoundingBox fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys + .head.NumEdges) * 4 + .head.NumPolys * 24 + .head.mirex_h * 100 + .head.NumGroups * 56 + 4, .BoundingBox
        WriteNormalIndex fileNumber, &H81 + (.head.NumVerts + .head.NumNormals) * 12 + .head.NumTexCs * 8 + (.head.NumVerts + .head.NumPolys + .head.NumEdges) * 4 + .head.NumPolys * 24 + .head.mirex_h * 100 + .head.NumGroups * 56 + 4 + 24, .NormalIndex
        .ResizeX = 1
        .ResizeY = 1
        .ResizeZ = 1
        .RotateAlpha = 0
        .RotateBeta = 0
        .RotateGamma = 0
        .RepositionX = 0
        .RepositionY = 0
        .RepositionZ = 0
    End With
    
    Close fileNumber
    Exit Sub
ErrHandWrite:
    'Debug.Print "Error writing P file!!!"
    MsgBox "Error writing P file!!!", vbOKOnly, "Error writing"
End Sub
Public Sub MergePModels(ByRef p1 As PModel, ByRef p2 As PModel)
    On Error GoTo hand
    With p1
1:        MergeVerts .Verts, p2.Verts
'2:        MergeNormals .Normals, p2.Normals
        If .head.NumTexCs = 0 Then
            p1.TexCoords = p2.TexCoords
        Else
            If p2.head.NumTexCs > 0 Then
3:              MergeTexCoords .TexCoords, p2.TexCoords
            End If
        End If
4:        MergeVColors .vcolors, p2.vcolors
5:        MergePColors .PColors, p2.PColors
        'MergeEdges .Edges, p2.Edges
6:        MergePolygons .polys, p2.polys
7:        MergeHundrets .hundrets, p2.hundrets
8:        MergeGroups .Groups, p2.Groups
9:        MergeBoundingBox .BoundingBox, p2.BoundingBox
'10:        MergeNormalIndex .NormalIndex, p2.NormalIndex
11:        MergeHeader .head, p2.head
    End With
    ComputeNormals p1
    ComputeEdges p1
    
    CheckModelConsistency p1
    Exit Sub
hand:
    MsgBox "Merging " + p1.fileName + " with " + p2.fileName + "!!!" + Str$(Erl), vbOKOnly, "Error merging"
End Sub
Sub CreateDListsFromPModel(ByRef obj As PModel)
    Dim gi As Integer
    
    With obj
        For gi = 0 To .head.NumGroups - 1
            CreateDListFromPGroup .Groups(gi), .polys, .Verts, .vcolors, .Normals, .TexCoords, .hundrets(gi)
        Next gi
    End With
End Sub
Sub FreePModelResources(ByRef obj As PModel)
    Dim gi As Integer
    
    With obj
        For gi = 0 To .head.NumGroups - 1
            FreeGroupResources .Groups(gi)
        Next gi
    End With
End Sub
Sub ComputePModelBoundingBox(ByRef obj As PModel, ByRef p_min As Point3D, ByRef p_max As Point3D)
    Dim p_min_aux As Point3D
    Dim p_max_aux As Point3D
    Dim MV_matrix(16) As Double
    
    glMatrixMode GL_MODELVIEW
    glPushMatrix
    glLoadIdentity
    With obj.BoundingBox
        p_min_aux.x = .min_x
        p_min_aux.y = .min_y
        p_min_aux.z = .min_z
    
        p_max_aux.x = .max_x
        p_max_aux.y = .max_y
        p_max_aux.z = .max_z
    End With
    
    With obj
        ConcatenateCameraModelView .RepositionX, .RepositionY, .RepositionZ, _
            .RotateAlpha, .RotateBeta, .RotateGamma, .ResizeX, .ResizeY, .ResizeZ
    End With
    
    glGetDoublev GL_MODELVIEW_MATRIX, MV_matrix(0)
            
    ComputeTransformedBoxBoundingBox MV_matrix, p_min_aux, p_max_aux, _
        p_min, p_max
    
    glMatrixMode GL_MODELVIEW
    glPopMatrix
End Sub
Sub SetCameraPModel(ByRef obj As PModel, ByVal cx As Single, ByVal cy As Single, ByVal CZ As Single, ByVal alpha As Single, ByVal Beta As Single, ByVal Gamma As Single, ByVal redX As Single, ByVal redY As Single, ByVal redZ As Single)
    Dim p_min As Point3D
    Dim p_max As Point3D
    Dim width As Integer
    Dim height As Integer
    Dim center_model As Point3D
    Dim model_radius As Single
    Dim distance_origin As Single
    Dim origin As Point3D
    Dim distance_radius As Single
    Dim scene_radius As Single
    Dim vp(4) As Long
    Dim rot_mat(16) As Double
    
    ComputePModelBoundingBox obj, p_min, p_max
    
    glGetIntegerv GL_VIEWPORT, vp(0)
    width = vp(2)
    height = vp(3)
    
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    center_model.x = (p_min.x + p_max.x) / 2
    center_model.y = (p_min.y + p_max.y) / 2
    center_model.z = (p_min.z + p_max.z) / 2
    origin.x = 0
    origin.y = 0
    origin.z = 0
    model_radius = CalculateDistance(p_min, p_max) / 2
    distance_origin = CalculateDistance(center_model, origin)
    scene_radius = model_radius + distance_origin
    gluPerspective 60, width / height, max(0.1, -CZ - scene_radius), max(0.1, -CZ + scene_radius)
    
    SetCameraModelView cx, cy, CZ, alpha, Beta, Gamma, redX, redY, redZ
End Sub
Sub DrawPModelDLists(ByRef obj As PModel, ByRef tex_ids() As Long)
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    
   ' glMatrixMode GL_MODELVIEW
   ' glPushMatrix
   ' With obj
   '     glScalef .ResizeX, .ResizeY, .ResizeZ
   '     glRotatef .RotateAlpha, 1, 0, 0
   '     glRotatef .RotateBeta, 0, 1, 0
   '     glRotatef .RotateGamma, 0, 0, 1
   '     glTranslatef .RepositionX, .RepositionY, .RepositionZ
   ' End With
    
    glShadeModel GL_SMOOTH
    glPolygonMode GL_FRONT, GL_FILL
    glPolygonMode GL_BACK, GL_FILL
    glEnable GL_COLOR_MATERIAL

    For gi = 0 To obj.head.NumGroups - 1
        If obj.Groups(gi).texFlag = 1 And tex_ids(0) > 0 Then
            ''Debug.Print "Is Texture?", glIsTexture(tex_ids(obj.Groups(gi).texID)) = GL_TRUE, tex_ids(obj.Groups(gi).texID)
            If (obj.Groups(gi).texID <= UBound(tex_ids)) Then
                If glIsTexture(tex_ids(obj.Groups(gi).texID)) = GL_TRUE Then
                    glEnable GL_TEXTURE_2D
    
                    glBindTexture GL_TEXTURE_2D, tex_ids(obj.Groups(gi).texID)
                    glTexParameterf GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
                    glTexParameterf GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
                    ''Debug.Print "Binding Texture to mesh...", glGetError
                Else
                    ''Debug.Print "Not a texture ", tex_ids(obj.Groups(gi).texID)
                End If
            End If
        End If
        With obj
            DrawGroupDList .Groups(gi)
            glDisable GL_TEXTURE_2D
        End With
    Next gi
    'glPopMatrix
End Sub

Sub DrawPModelBoundingBox(ByRef obj As PModel)
    glBegin GL_LINES
        glDisable GL_DEPTH_TEST
        With obj.BoundingBox
            DrawBox .max_x, .max_y, .max_z, .min_x, .min_y, .min_z, 1, 1, 0
        End With
        glEnable GL_DEPTH_TEST
    glEnd
    'glPopMatrix
End Sub
Sub DrawPModel(ByRef obj As PModel, ByRef tex_ids() As Long, ByVal HideHiddenGroupsQ As Boolean)
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    Dim set_v_textured As Boolean
    Dim v_textured As Boolean
    Dim set_v_linearfilter As Boolean
    Dim v_linearfilter As Boolean
    
    Dim TexEnabled As Boolean
    
    TexEnabled = (glIsEnabled(GL_TEXTURE_2D) = GL_TRUE)
    
    glShadeModel GL_SMOOTH
    glEnable GL_COLOR_MATERIAL

    For gi = 0 To obj.head.NumGroups - 1
        'Set the render states acording to the hundrets information
        'V_WIREFRAME
        If Not ((obj.hundrets(gi).field_8 And &H1) = 0) Then
            If Not ((obj.hundrets(gi).field_C And &H1) = 0) Then
                glPolygonMode GL_FRONT, GL_LINE
                glPolygonMode GL_BACK, GL_LINE
            Else
                glPolygonMode GL_FRONT, GL_FILL
                glPolygonMode GL_BACK, GL_FILL
            End If
        End If
        
        'V_TEXTRED
        set_v_textured = Not ((obj.hundrets(gi).field_8 And &H2) = 0)
        v_textured = Not ((obj.hundrets(gi).field_C And &H2) = 0)
        
        'V_LINEARFILTER
        set_v_linearfilter = Not ((obj.hundrets(gi).field_8 And &H4) = 0)
        v_linearfilter = Not ((obj.hundrets(gi).field_C And &H4) = 0)
        
        'V_NOCULL
        If Not ((obj.hundrets(gi).field_8 And &H4000&) = 0) Then
            If Not ((obj.hundrets(gi).field_C And &H4000&) = 0) Then
                glDisable GL_CULL_FACE
            Else
                glEnable GL_CULL_FACE
            End If
        End If
        
        'V_CULLFACE
        If Not ((obj.hundrets(gi).field_8 And &H2000&) = 0) Then
            If Not ((obj.hundrets(gi).field_C And &H2000&) = 0) Then
                glCullFace GL_FRONT
            Else
                glCullFace GL_BACK
            End If
        End If
        
        'Now let's set the blending state
        Select Case obj.hundrets(gi).blend_mode
            Case 0:
                'Average
                glEnable GL_BLEND
                GL_Ext.glBlendEquation GL_FUNC_ADD
                If (TexEnabled And Not (set_v_textured And Not v_textured)) Or _
                   (set_v_textured And v_textured) Then
                    glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
                Else
                    glBlendFunc GL_SRC_ALPHA, GL_SRC_ALPHA
                End If
            Case 1:
                'Additive
                glEnable GL_BLEND
                GL_Ext.glBlendEquation GL_FUNC_ADD
                glBlendFunc GL_ONE, GL_ONE
            Case 2:
                'Subtractive
                glEnable GL_BLEND
                GL_Ext.glBlendEquation GL_FUNC_REVERSE_SUBTRACT
                glBlendFunc GL_ONE, GL_ONE
            Case 3:
                'Unknown, let's disable blending
                glDisable GL_BLEND
            Case 4:
                'No blending
                glDisable GL_BLEND
                If Not ((obj.hundrets(gi).field_8 And &H400&) = 0) Then
                    If Not ((obj.hundrets(gi).field_C And &H400&) = 0) Then
                        glEnable GL_BLEND
                        GL_Ext.glBlendEquation GL_FUNC_ADD
                        glBlendFunc GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
                    End If
                End If
        End Select
        
        
        If obj.Groups(gi).texFlag = 1 And tex_ids(0) > 0 And v_textured Then
            ''Debug.Print "Is Texture?", glIsTexture(tex_ids(obj.Groups(gi).texID)) = GL_TRUE, tex_ids(obj.Groups(gi).texID)
            If (obj.Groups(gi).texID <= UBound(tex_ids)) Then
                If glIsTexture(tex_ids(obj.Groups(gi).texID)) = GL_TRUE Then
                    If set_v_textured Then
                        If v_textured Then
                            glEnable GL_TEXTURE_2D
                        Else
                            glDisable GL_TEXTURE_2D
                        End If
                    End If
    
                    glBindTexture GL_TEXTURE_2D, tex_ids(obj.Groups(gi).texID)
                    
                    If set_v_linearfilter Then
                        If v_linearfilter Then
                            glTexParameterf GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
                            glTexParameterf GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
                        Else
                            glTexParameterf GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST
                            glTexParameterf GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
                        End If
                    End If
                Else
                    ''Debug.Print "Not a texture ", tex_ids(obj.Groups(gi).texID)
                End If
            End If
        End If
        With obj
            DrawGroup .Groups(gi), .polys, .Verts, .vcolors, .Normals, .TexCoords, .hundrets(gi), HideHiddenGroupsQ
            glDisable GL_TEXTURE_2D
        End With
    Next gi
    'glPopMatrix
End Sub
Sub DrawPModelPolys(ByRef obj As PModel)
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    
    glShadeModel GL_FLAT
    glPolygonMode GL_FRONT, GL_LINE
    glPolygonMode GL_BACK, GL_FILL
    glEnable GL_COLOR_MATERIAL
    
    
    For gi = 0 To obj.head.NumGroups - 1
        For PI = obj.Groups(gi).offpoly To obj.Groups(gi).offpoly + obj.Groups(gi).numPoly - 1
            With obj.PColors(PI)
                glColor4f .r / 255, .g / 255, .B / 255, .a / 255
                glColorMaterial GL_FRONT_AND_BACK, GL_AMBIENT_AND_DIFFUSE
            End With
            glBegin GL_TRIANGLES
                For vi = 0 To 2
                    With obj.Normals(obj.polys(PI).Normals(vi))
                        glNormal3f .x, .y, .z
                    End With
                    With obj.Verts(obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert)
                        glVertex3f .x, .y, .z
                    End With
                Next vi
            glEnd
        Next PI
    Next gi
    'glPopMatrix
End Sub
Sub DrawPModelMesh(ByRef obj As PModel)
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    
    'glMatrixMode GL_MODELVIEW
    'glPushMatrix
    'With obj
    '    glScalef .ResizeX, .ResizeY, .ResizeZ
    '    glRotatef .RotateAlpha, 1, 0, 0
    '    glRotatef .RotateBeta, 0, 1, 0
    '    glRotatef .RotateGamma, 0, 0, 1
    '    glTranslatef .RepositionX, .RepositionY, .RepositionZ
    'End With
    
    glPolygonMode GL_FRONT, GL_LINE
    glPolygonMode GL_BACK, GL_LINE
    glColor3f 0, 0, 0
    
    For gi = 0 To obj.head.NumGroups - 1
        For PI = obj.Groups(gi).offpoly To obj.Groups(gi).offpoly + obj.Groups(gi).numPoly - 1
            glBegin GL_TRIANGLES
                For vi = 0 To 2
                    With obj.Verts(obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert)
                        glVertex3f .x, .y, .z
                    End With
                Next vi
            glEnd
        Next PI
    Next gi
    'glPopMatrix
End Sub
Sub DrawVert(ByRef object As PModel, ByVal vi As Integer)
    'glMatrixMode GL_MODELVIEW
    'glPushMatrix
    'With obj
    '    glScalef .ResizeX, .ResizeY, .ResizeZ
    '    glRotatef .RotateAlpha, 1, 0, 0
    '    glRotatef .RotateBeta, 0, 1, 0
    '    glRotatef .RotateGamma, 0, 0, 1
    '    glTranslatef .RepositionX, .RepositionY, .RepositionZ
    'End With
    DrawVertT object.Verts, vi
    'glPopMatrix
End Sub
'----------------------------------------------------------------------------------------------------
'=============================================REPAIRING==============================================
'----------------------------------------------------------------------------------------------------
Sub ComputeBoundingBox(ByRef obj As PModel)
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    
    With obj.BoundingBox
        .max_x = -INFINITY_SINGLE
        .max_y = -INFINITY_SINGLE
        .max_z = -INFINITY_SINGLE
        .min_x = INFINITY_SINGLE
        .min_y = INFINITY_SINGLE
        .min_z = INFINITY_SINGLE
    End With
    
    For gi = 0 To obj.head.NumGroups - 1
        For PI = obj.Groups(gi).offpoly To obj.Groups(gi).offpoly + obj.Groups(gi).numPoly - 1
            For vi = 0 To 2
                With obj.Verts(obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert)
                    If .x > obj.BoundingBox.max_x Then obj.BoundingBox.max_x = .x
                    If .y > obj.BoundingBox.max_y Then obj.BoundingBox.max_y = .y
                    If .z > obj.BoundingBox.max_z Then obj.BoundingBox.max_z = .z
                    
                    If .x < obj.BoundingBox.min_x Then obj.BoundingBox.min_x = .x
                    If .y < obj.BoundingBox.min_y Then obj.BoundingBox.min_y = .y
                    If .z < obj.BoundingBox.min_z Then obj.BoundingBox.min_z = .z
                End With
            Next vi
        Next PI
    Next gi
    
    obj.diameter = Sqr(obj.BoundingBox.max_x ^ 2 + obj.BoundingBox.max_y ^ 2 + obj.BoundingBox.max_z ^ 2 + _
                        obj.BoundingBox.min_x ^ 2 + obj.BoundingBox.min_y ^ 2 + obj.BoundingBox.min_z ^ 2)
End Sub
Sub ComputeCurrentBoundingBox(ByRef obj As PModel)
    Dim p_temp As Point3D
    
    ComputeBoundingBox obj
    
    With obj.BoundingBox
        p_temp.x = .max_x
        p_temp.y = .max_y
        p_temp.z = .max_z
    End With
    p_temp = GetEyeSpaceCoords(p_temp)
    With obj.BoundingBox
        .max_x = p_temp.x
        .max_y = p_temp.y
        .max_z = p_temp.z
    End With
    With obj.BoundingBox
        p_temp.x = .min_x
        p_temp.y = .min_y
        p_temp.z = .min_z
    End With
    p_temp = GetEyeSpaceCoords(p_temp)
    With obj.BoundingBox
        .min_x = p_temp.x
        .min_y = p_temp.y
        .min_z = p_temp.z
    End With
End Sub
Sub ComputeNormals(ByRef obj As PModel)
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    Dim vi2 As Integer
    
    Dim temp_norm As Point3D
    
    ReDim obj.Normals(obj.head.NumVerts)
    ReDim obj.NormalIndex(obj.head.NumVerts)
    
    obj.head.NumNormals = obj.head.NumVerts
    obj.head.NumNormInds = obj.head.NumVerts
    
    Dim sum_norms() As Point3D
    ReDim sum_norms(obj.head.NumVerts)
    Dim sum_temp As Point3D
    
    Dim polys_per_vert() As Integer
    ReDim polys_per_vert(obj.head.NumVerts)
    Dim polys_temp As Integer
    
    With obj
        'This should never happen. What the hell is going on?!
        For PI = 0 To .head.NumPolys - 1
            If .polys(PI).Verts(0) < 0 Then _
                .polys(PI).Verts(0) = 0
            If .polys(PI).Verts(1) < 0 Then _
                .polys(PI).Verts(1) = 0
            If .polys(PI).Verts(2) < 0 Then _
                .polys(PI).Verts(2) = 0
        Next PI
    End With

    For gi = 0 To obj.head.NumGroups - 1
        For PI = obj.Groups(gi).offpoly To obj.Groups(gi).offpoly + obj.Groups(gi).numPoly - 1
            With obj
                temp_norm = CalculateNormal(.Verts(.polys(PI).Verts(0) + .Groups(gi).offvert), _
                                            .Verts(.polys(PI).Verts(1) + .Groups(gi).offvert), _
                                            .Verts(.polys(PI).Verts(2) + .Groups(gi).offvert))
            End With
            For vi = 0 To 2
                With sum_norms(obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert)
                    .x = .x + temp_norm.x
                    .y = .y + temp_norm.y
                    .z = .z + temp_norm.z
                End With
                polys_per_vert(obj.polys(PI).Verts(vi) _
                               + obj.Groups(gi).offvert) = 1 + polys_per_vert(obj.polys(PI).Verts(vi) _
                                                                              + obj.Groups(gi).offvert)
                
                obj.polys(PI).Normals(vi) = obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert
            Next vi
        Next PI
    Next gi
    
    'If False Then
    For vi = 0 To obj.head.NumVerts - 1
        If polys_per_vert(vi) > 0 Then
            For vi2 = vi + 1 To obj.head.NumVerts - 1
                If ComparePoints3D(obj.Verts(vi), obj.Verts(vi2)) Then
                    With sum_norms(vi)
                        .x = .x + sum_norms(vi2).x
                        .y = .y + sum_norms(vi2).y
                        .z = .z + sum_norms(vi2).z
                    End With
                    
                    sum_norms(vi2) = sum_norms(vi)

                    polys_per_vert(vi) = polys_per_vert(vi) + polys_per_vert(vi2)
                    polys_per_vert(vi2) = -polys_per_vert(vi)
                End If
            Next vi2
        Else
            For vi2 = vi + 1 To obj.head.NumVerts - 1
                If ComparePoints3D(obj.Verts(vi), obj.Verts(vi2)) Then
                    sum_norms(vi) = sum_norms(vi2)

                    polys_per_vert(vi) = -polys_per_vert(vi2)
                End If
            Next vi2
        End If
        polys_per_vert(vi) = Abs(polys_per_vert(vi))
    Next vi
    'End If
    
    For vi = 0 To obj.head.NumVerts - 1
        If polys_per_vert(vi) > 0 Then
            With sum_norms(vi)
                .x = -.x / polys_per_vert(vi)
                .y = -.y / polys_per_vert(vi)
                .z = -.z / polys_per_vert(vi)
            End With
        Else
            With sum_norms(vi)
                ''Debug.Print vi, polys_per_vert(vi)
                .x = 0
                .y = 0
                .z = 0
            End With
        End If
        
        obj.Normals(vi) = Normalize(sum_norms(vi))
        obj.NormalIndex(vi) = vi
    Next vi
End Sub


Sub DisableNormals(ByRef obj As PModel)
    Dim vi As Integer
    Dim PI As Integer
    Dim gi As Integer
    Dim nii As Integer
    
    ReDim obj.Normals(0)
    obj.NormalIndex(0) = 0
    
    For PI = 0 To obj.head.NumPolys - 1
        With obj.polys(PI)
            For vi = 0 To 2
                .Normals(vi) = 0
            Next vi
        End With
    Next PI
    
    For gi = 0 To obj.head.NumGroups - 1
        With obj.Groups(gi)
            If .polyType = 2 Then .polyType = 3
        End With
    Next gi
    
    For nii = 0 To obj.head.NumNormInds - 1
        obj.NormalIndex(nii) = 0
    Next nii
    
    obj.head.NumNormals = 0
End Sub
Sub ComputeEdges(ByRef obj As PModel)
    Dim gi As Integer
    Dim PI As Integer
    Dim ei As Integer
    Dim vi As Integer
    
    ReDim obj.Edges(obj.head.NumPolys * 3)
    
    Dim num_edges As Integer
    Dim found As Boolean
    
    For gi = 0 To obj.head.NumGroups - 1
    '    obj.Groups(gi).offEdge = num_edges
    '    For pi = obj.Groups(gi).offPoly To obj.Groups(gi).offPoly + obj.Groups(gi).numPoly - 1
    '        For vi = 0 To 2
    '            found = False
    '            For ei = 0 To num_edges - 1
    '                With obj.Edges(ei)
    '                    If (.Verts(0) = obj.Polys(pi).Verts(vi) And _
    '                        .Verts(1) = obj.Polys(pi).Verts((vi + 1) Mod 3)) Or _
    '                       (.Verts(1) = obj.Polys(pi).Verts(vi) And _
    '                        .Verts(0) = obj.Polys(pi).Verts((vi + 1) Mod 3)) Then
    '                        found = True
    '                        Exit For
    '                    End If
    '                End With
    '            Next ei
    '
    '            If Not found Then
    '                With obj.Edges(num_edges)
    '                    .Verts(0) = obj.Polys(pi).Verts(vi)
    '                    .Verts(1) = obj.Polys(pi).Verts((vi + 1) Mod 3)
    '                End With
    '                obj.Polys(pi).Edges(vi) = num_edges - obj.Groups(gi).offEdge
    '                num_edges = num_edges + 1
    '            Else
    '                obj.Polys(pi).Edges(vi) = ei - obj.Groups(gi).offEdge
    '            End If
    '        Next vi
    '    Next pi
    '
        obj.Groups(gi).numEdge = obj.Groups(gi).numPoly * 3 'num_edges - obj.Groups(gi).offEdge
    Next gi
    
    obj.head.NumEdges = obj.head.NumPolys * 3 'num_edges
End Sub
Function ComputePolyColor(ByRef obj As PModel, ByVal PI As Integer) As color
    Dim vi As Integer
    Dim Group As Integer
    
    Dim temp_a As Integer
    Dim temp_r As Integer
    Dim temp_g As Integer
    Dim temp_b As Integer
    
    Group = GetPolygonGroup(obj.Groups, PI)
    
    For vi = 0 To 2
        With obj.vcolors(obj.polys(PI).Verts(vi) + obj.Groups(Group).offvert)
            temp_a = temp_a + .a
            temp_r = temp_r + .r
            temp_g = temp_g + .g
            temp_b = temp_b + .B
        End With
    Next vi
    
    With ComputePolyColor
        .a = temp_a / 3
        .r = temp_r / 3
        .g = temp_g / 3
        .B = temp_b / 3
    End With
End Function
Sub ComputePColors(ByRef obj As PModel)
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    
    Dim temp_a As Integer
    Dim temp_r As Integer
    Dim temp_g As Integer
    Dim temp_b As Integer
    
    ReDim obj.PColors(obj.head.NumPolys - 1)
    
    For gi = 0 To obj.head.NumGroups - 1
        For PI = obj.Groups(gi).offpoly To obj.Groups(gi).offpoly + obj.Groups(gi).numPoly - 1
            temp_r = 0
            temp_g = 0
            temp_b = 0
            For vi = 0 To 2
                With obj.vcolors(obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert)
                    temp_r = temp_r + .r
                    temp_g = temp_g + .g
                    temp_b = temp_b + .B
                End With
            Next vi
            With obj.PColors(PI)
                .a = 255
                .r = temp_r / 3
                .g = temp_g / 3
                .B = temp_b / 3
            End With
        Next PI
    Next gi
End Sub
Sub KillEmptyGroups(ByRef obj As PModel)
    Dim gi As Integer
    
    gi = 0
    While gi < obj.head.NumGroups - 1
        If (obj.Groups(gi).numvert = 0) Then
           RemoveGroup obj, gi
        Else
            gi = gi + 1
        End If
    Wend
End Sub
Sub KillUnusedVertices(ByRef obj As PModel)
'------------------------------�WARNINGS!------------------------------
'-------*Causes the Normals to be inconsistent (call ComputeNormals).--
    Dim gi As Long
    Dim gi2 As Long
    Dim PI As Long
    Dim vi As Long
    Dim vi2 As Long
    Dim vit As Long
    Dim tci As Long
    Dim tci_global As Long
    
    Dim verts_usage() As Long
    
    ReDim verts_usage(obj.head.NumVerts - 1)
    
    For vi = 0 To obj.head.NumVerts - 1
        verts_usage(vi) = 0
    Next vi
    
    With obj
        For gi = 0 To .head.NumGroups - 1
            For PI = .Groups(gi).offpoly To .Groups(gi).offpoly + .Groups(gi).numPoly - 1
                For vi = 0 To 2
                    verts_usage(.polys(PI).Verts(vi) + .Groups(gi).offvert) = 1 + _
                        verts_usage(.polys(PI).Verts(vi) + .Groups(gi).offvert)
                Next vi
            Next PI
        Next gi
        
        vi = 0
        vit = 0
        tci_global = 0
        For gi = 0 To .head.NumGroups - 1
            While vi < .Groups(gi).offvert + .Groups(gi).numvert
                If verts_usage(vit) = 0 Then
                    'If the vertex is unused, let's destory it
                    For vi2 = vi To .head.NumVerts - 2
                        .Verts(vi2) = .Verts(vi2 + 1)
                        .vcolors(vi2) = .vcolors(vi2 + 1)
                    Next vi2
                    
                    If .Groups(gi).texFlag = 1 Then
                        For tci = tci_global To .head.NumTexCs - 2
                            .TexCoords(tci) = .TexCoords(tci + 1)
                        Next tci
                        
                        .head.NumTexCs = .head.NumTexCs - 1
                        ReDim Preserve .TexCoords(.head.NumTexCs - 1)
                    End If
                    
                    .head.NumVerts = .head.NumVerts - 1
                    ReDim Preserve .Verts(.head.NumVerts - 1)
                    ReDim Preserve .vcolors(.head.NumVerts - 1)
                    
                    For PI = .Groups(gi).offpoly To .Groups(gi).offpoly + .Groups(gi).numPoly - 1
                        For vi2 = 0 To 2
                            If .polys(PI).Verts(vi2) > vi - .Groups(gi).offvert Then
                                .polys(PI).Verts(vi2) = .polys(PI).Verts(vi2) - 1
                            End If
                        Next vi2
                    Next PI
                    
                    If gi < .head.NumGroups - 1 Then
                        For gi2 = gi + 1 To .head.NumGroups - 1
                            .Groups(gi2).offvert = .Groups(gi2).offvert - 1
                            If .Groups(gi).texFlag = 1 And _
                               .Groups(gi2).offTex > 0 Then _
                                .Groups(gi2).offTex = .Groups(gi2).offTex - 1
                        Next gi2
                    End If
                    .Groups(gi).numvert = .Groups(gi).numvert - 1
                Else
                    vi = vi + 1
                    If .Groups(gi).texFlag = 1 Then _
                        tci_global = tci_global + 1
                End If
                vit = vit + 1
            Wend
        Next gi
    End With
End Sub
Sub KillCrappyPolygons(ByRef obj As PModel) 'Kill degenerated polygons (points and lines)
    Dim gi As Integer
    Dim PI As Integer
    
    For gi = 0 To obj.head.NumGroups - 1
        PI = obj.Groups(gi).offpoly
        While PI < obj.Groups(gi).offpoly + obj.Groups(gi).numPoly - 1
            If ComparePoints3D(obj.Verts(obj.polys(PI).Verts(0) + obj.Groups(gi).offvert), obj.Verts(obj.polys(PI).Verts(1) + obj.Groups(gi).offvert)) And _
               ComparePoints3D(obj.Verts(obj.polys(PI).Verts(0) + obj.Groups(gi).offvert), obj.Verts(obj.polys(PI).Verts(2) + obj.Groups(gi).offvert)) Then
                RemovePolygon obj, PI
            Else
                PI = PI + 1
            End If
        Wend
    Next gi
End Sub
'----------------------------------------------------------------------------------------------------
'=============================================SELECTORS==============================================
'----------------------------------------------------------------------------------------------------
Function GetClosestPolygon(ByRef obj As PModel, ByVal px As Integer, ByVal py As Integer, ByVal DIST As Single) As Integer
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    
    Dim min_z As Single
    Dim spi As Integer
    Dim nPolys As Integer
    
    Dim vp(4) As Long
    Dim P_matrix(16) As Double
    
    Dim Sel_BUFF() As Long
    ReDim Sel_BUFF(obj.head.NumPolys * 4)
    
    Dim width As Integer
    Dim height As Integer
    
    glSelectBuffer obj.head.NumPolys * 4, Sel_BUFF(0)
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
    glMultMatrixd P_matrix(0)
    
    For gi = 0 To obj.head.NumGroups - 1
        If Not obj.Groups(gi).HiddenQ Then
            For PI = obj.Groups(gi).offpoly To obj.Groups(gi).offpoly + obj.Groups(gi).numPoly - 1
                glPushName PI
                glBegin GL_TRIANGLES
                    For vi = 0 To 2
                        With obj.Verts(obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert)
                            glVertex3f .x, .y, .z
                        End With
                    Next vi
                glEnd
                glPopName
            Next PI
        End If
    Next gi
    
    nPolys = glRenderMode(GL_RENDER)
    GetClosestPolygon = -1
    min_z = -1
    
    For PI = 0 To nPolys - 1
        If CompareLongs(min_z, Sel_BUFF(PI * 4 + 1)) Or _
          (Sel_BUFF(PI * 4 + 1) = min_z) Then 'And _
           obj.Groups(GetPolygonGroup(obj.Groups, Sel_BUFF(PI * 4 + 3))).texFlag <> 1) Then
            min_z = Sel_BUFF(PI * 4 + 1)
            GetClosestPolygon = Sel_BUFF(PI * 4 + 3)
        End If
    Next PI
    glMatrixMode GL_PROJECTION
    glPopMatrix
End Function
Function GetClosestVertex(ByRef obj As PModel, ByVal px As Integer, ByVal py As Integer, ByVal DIST0 As Single) As Integer
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    Dim vi2 As Integer
    Dim width As Integer
    Dim height As Integer
    Dim viewport(4) As Long
    
    Dim p As Point3D
    Dim vp As Point3D
    
    Dim DIST(2) As Single
    Dim min_dist As Single
    
    p.x = px
    p.y = py
    p.z = 0
    
    glGetIntegerv GL_VIEWPORT, viewport(0)
    width = viewport(2)
    height = viewport(3)
    
    PI = GetClosestPolygon(obj, px, py, DIST0)
    
    If PI > -1 Then
        
        gi = GetPolygonGroup(obj.Groups, PI)
        
        p.y = height - py
        For vi = 0 To 2
            vi2 = obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert
            vp = GetVertexProjectedCoords(obj.Verts, vi2)
            DIST(vi) = CalculateDistance(vp, p)
        Next vi
        
        min_dist = DIST(0)
        GetClosestVertex = obj.polys(PI).Verts(0) + obj.Groups(gi).offvert
        
        For vi = 1 To 2
            If DIST(vi) < min_dist Then
                min_dist = DIST(vi)
                GetClosestVertex = obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert
            End If
        Next vi
    Else
        GetClosestVertex = -1
    End If
End Function
Function GetClosestEdge(ByRef obj As PModel, ByVal p_index As Integer, ByVal px As Integer, ByVal py As Integer, ByRef alpha As Single) As Integer
    Dim p_temp As Point3D
    
    Dim p1_proj As Point3D
    Dim p2_proj As Point3D
    Dim p3_proj As Point3D
    
    Dim p1 As Point3D
    Dim p2 As Point3D
    Dim p3 As Point3D
    
    Dim d1 As Single
    Dim d2 As Single
    Dim d3 As Single
    
    Dim viewport(4) As Long
    Dim width As Integer
    Dim height As Integer
    Dim offVerts As Integer
    
    glGetIntegerv GL_VIEWPORT, viewport(0)
    width = viewport(2)
    height = viewport(3)
    
    p_temp.x = px
    p_temp.y = height - py
    p_temp.z = 0
    
    offVerts = obj.Groups(GetPolygonGroup(obj.Groups, p_index)).offvert
    
    'glMatrixMode GL_MODELVIEW
    'glPushMatrix
    'With obj
    '    glScalef .ResizeX, .ResizeY, .ResizeZ
    '    glRotatef .RotateAlpha, 1, 0, 0
    '    glRotatef .RotateBeta, 0, 1, 0
    '    glRotatef .RotateGamma, 0, 0, 1
    '    glTranslatef .RepositionX, .RepositionY, .RepositionZ
    'End With
    
    With obj.polys(p_index)
        p1_proj = GetVertexProjectedCoords(obj.Verts, .Verts(0) + offVerts)
        p2_proj = GetVertexProjectedCoords(obj.Verts, .Verts(1) + offVerts)
        p3_proj = GetVertexProjectedCoords(obj.Verts, .Verts(2) + offVerts)
        
        p1 = CalculatePoint2LineProjection(p_temp, p1_proj, p2_proj)
        p2 = CalculatePoint2LineProjection(p_temp, p2_proj, p3_proj)
        p3 = CalculatePoint2LineProjection(p_temp, p3_proj, p1_proj)
    
        d1 = CalculateDistance(p_temp, p1)
        d2 = CalculateDistance(p_temp, p2)
        d3 = CalculateDistance(p_temp, p3)
        
        If d1 > d2 Then
            If d2 > d3 Then
                GetClosestEdge = 2
                alpha = CalculatePoint2LineProjectionPosition(p_temp, p3_proj, p1_proj)
            Else
                GetClosestEdge = 1
                alpha = CalculatePoint2LineProjectionPosition(p_temp, p2_proj, p3_proj)
            End If
        Else
            If d1 > d3 Then
                GetClosestEdge = 2
                alpha = CalculatePoint2LineProjectionPosition(p_temp, p3_proj, p1_proj)
            Else
                GetClosestEdge = 0
                alpha = CalculatePoint2LineProjectionPosition(p_temp, p1_proj, p2_proj)
            End If
        End If
    End With
    'glMatrixMode GL_MODELVIEW
    'glPopMatrix
End Function
'Find the first edge between v1 and v2 (poly and edge id)
Public Function FindNextAdjacentPolyEdge(ByRef obj As PModel, ByRef v1 As Point3D, ByRef v2 As Point3D, ByRef p_index As Integer, ByRef e_index As Integer) As Boolean
    Dim PI As Integer
    Dim gi As Integer
    Dim found As Boolean
    Dim offvert As Long
    
    found = False
    
    For gi = 0 To obj.head.NumGroups - 1
        If Not obj.Groups(gi).HiddenQ Then
            offvert = obj.Groups(gi).offvert
            For PI = obj.Groups(gi).offpoly To obj.Groups(gi).offpoly + obj.Groups(gi).numPoly - 1
                With obj.polys(PI)
                    If (ComparePoints3D(obj.Verts(offvert + .Verts(0)), v1) And _
                        ComparePoints3D(obj.Verts(offvert + .Verts(1)), v2)) Or _
                       (ComparePoints3D(obj.Verts(offvert + .Verts(0)), v2) And _
                        ComparePoints3D(obj.Verts(offvert + .Verts(1)), v1)) Then
                        p_index = PI
                        e_index = 0
                        found = True
                        Exit For
                    Else
                        If (ComparePoints3D(obj.Verts(offvert + .Verts(1)), v1) And _
                            ComparePoints3D(obj.Verts(offvert + .Verts(2)), v2)) Or _
                           (ComparePoints3D(obj.Verts(offvert + .Verts(1)), v2) And _
                            ComparePoints3D(obj.Verts(offvert + .Verts(2)), v1)) Then
                            p_index = PI
                            e_index = 1
                            found = True
                            Exit For
                        Else
                            If (ComparePoints3D(obj.Verts(offvert + .Verts(2)), v1) And _
                                ComparePoints3D(obj.Verts(offvert + .Verts(0)), v2)) Or _
                               (ComparePoints3D(obj.Verts(offvert + .Verts(2)), v2) And _
                                ComparePoints3D(obj.Verts(offvert + .Verts(0)), v1)) Then
                                p_index = PI
                                e_index = 2
                                found = True
                                Exit For
                            End If
                        End If
                    End If
                End With
            Next PI
        End If
    Next gi
    
    FindNextAdjacentPolyEdge = found
End Function
'This version of the function find the next matching edge after the one given as parameter
Public Function FindNextAdjacentPolyEdgeForward(ByRef obj As PModel, ByRef v1 As Point3D, ByRef v2 As Point3D, ByRef g_index As Integer, ByRef p_index As Integer, ByRef e_index As Integer) As Boolean
    Dim PI As Integer
    Dim gi As Integer
    Dim found As Boolean
    Dim offvert As Long
    
    found = False
    
    PI = p_index + 1
    For g_index = g_index To obj.head.NumGroups - 1
        If Not obj.Groups(g_index).HiddenQ Then
            offvert = obj.Groups(g_index).offvert
            While PI < obj.Groups(g_index).offpoly + obj.Groups(g_index).numPoly
                With obj.polys(PI)
                    If (ComparePoints3D(obj.Verts(offvert + .Verts(0)), v1) And _
                        ComparePoints3D(obj.Verts(offvert + .Verts(1)), v2)) Or _
                       (ComparePoints3D(obj.Verts(offvert + .Verts(0)), v2) And _
                        ComparePoints3D(obj.Verts(offvert + .Verts(1)), v1)) Then
                        p_index = PI
                        e_index = 0
                        found = True
                        Exit For
                    Else
                        If (ComparePoints3D(obj.Verts(offvert + .Verts(1)), v1) And _
                            ComparePoints3D(obj.Verts(offvert + .Verts(2)), v2)) Or _
                           (ComparePoints3D(obj.Verts(offvert + .Verts(1)), v2) And _
                            ComparePoints3D(obj.Verts(offvert + .Verts(2)), v1)) Then
                            p_index = PI
                            e_index = 1
                            found = True
                            Exit For
                        Else
                            If (ComparePoints3D(obj.Verts(offvert + .Verts(2)), v1) And _
                                ComparePoints3D(obj.Verts(offvert + .Verts(0)), v2)) Or _
                               (ComparePoints3D(obj.Verts(offvert + .Verts(2)), v2) And _
                                ComparePoints3D(obj.Verts(offvert + .Verts(0)), v1)) Then
                                p_index = PI
                                e_index = 2
                                found = True
                                Exit For
                            End If
                        End If
                    End If
                End With
                PI = PI + 1
            Wend
        End If
    Next g_index
    
    FindNextAdjacentPolyEdgeForward = found
End Function
'---------------------------------------------------------------------------------------------------------
'--------------------------------------------------SETTERS------------------------------------------------
'---------------------------------------------------------------------------------------------------------
Sub AddGroup(ByRef obj As PModel, ByRef vertsV() As Point3D, ByRef facesV() As PPolygon, ByRef TexCoordsV() As Point2D, ByRef vcolorsV() As color, ByRef pcolorsV() As color)
'------------------- Warning! Causes the Normals to be inconsistent.------------------------------
'--------------------------------Must call ComputeNormals ----------------------------------------
    Dim gi As Integer
    Dim Group_index As Integer
    Dim offvert As Integer
    Dim offpoly As Integer
    Dim num_verts As Integer
    Dim num_polys As Integer
    Dim num_tex_coords As Integer
    
    num_verts = UBound(vertsV) + 1
    num_polys = UBound(facesV) + 1
    
    If SafeArrayGetDim(TexCoordsV) <> 0 Then
        num_tex_coords = UBound(TexCoordsV) + 1
    Else
        num_tex_coords = 0
    End If
    
    With obj
        If SafeArrayGetDim(.Groups) <> 0 Then
            Group_index = UBound(.Groups) + 1
        Else
            Group_index = 0
        End If
        
        ReDim Preserve .Groups(Group_index)
    End With
    
    With obj.Groups(Group_index)
        .polyType = IIf(num_tex_coords > 0, 2, 1)
        If SafeArrayGetDim(obj.polys) <> 0 Then
            .offpoly = UBound(obj.polys) + 1
        Else
            .offpoly = 0
        End If
        .numPoly = num_polys
        If SafeArrayGetDim(obj.Verts) <> 0 Then
            .offvert = UBound(obj.Verts) + 1
        Else
            .offvert = 0
        End If
        .numvert = num_verts
        .offEdge = 0
        .numEdge = 0
        .off1c = 0
        .off20 = 0
        .off24 = 0
        .off28 = 0
        If SafeArrayGetDim(obj.TexCoords) <> 0 Then
            .offTex = UBound(obj.TexCoords) + 1
        Else
            .offTex = 0
        End If
        .texFlag = IIf(num_tex_coords > 0, 1, 0)
        .texID = 0
        For gi = 0 To Group_index - 1
            If .texID <= obj.Groups(gi).texID Then _
                .texID = obj.Groups(gi).texID + 1
        Next gi
        .HiddenQ = False
    End With
    
    With obj.head
        .NumVerts = .NumVerts + num_verts
        .NumPolys = .NumPolys + num_polys
        .NumTexCs = .NumTexCs + num_tex_coords
        .NumGroups = .NumGroups + 1
        .mirex_h = .mirex_h + 1
        .mirex_g = 1
    End With
       
    With obj
        ReDim Preserve .Verts(.head.NumVerts - 1)
        CopyMemory .Verts(.head.NumVerts - num_verts), vertsV(0), CLng(num_verts) * 3 * 4
        ReDim Preserve .polys(.head.NumPolys - 1)
        CopyMemory .polys(.head.NumPolys - num_polys), facesV(0), CLng(num_polys) * 24
        If num_tex_coords > 0 Then
            ReDim Preserve .TexCoords(.head.NumTexCs - 1)
                'Debug.Print .head.NumTexCs - num_tex_coords
            CopyMemory .TexCoords(obj.Groups(Group_index).offTex), TexCoordsV(0), CLng(num_tex_coords) * 2 * 4
        End If
        ReDim Preserve .vcolors(.head.NumVerts - 1)
        CopyMemory .vcolors(.head.NumVerts - num_verts), vcolorsV(0), CLng(num_verts) * 4
        ReDim Preserve .PColors(.head.NumPolys - 1)
        CopyMemory .PColors(.head.NumPolys - num_polys), pcolorsV(0), CLng(num_polys) * 4
        ReDim Preserve .hundrets(.head.mirex_h - 1)
        FillHundrestsDefaultValues .hundrets(.head.mirex_h - 1)
    End With
End Sub
Function AddVertex(ByRef obj As PModel, ByVal Group As Integer, ByRef v As Point3D, ByRef vc As color) As Integer
'-------- Warning! Causes the Normals to be inconsistent if lights are disabled.------------------
'--------------------------------Must call ComputeNormals ----------------------------------------
    Dim gi As Integer
    Dim vi As Integer
    Dim ni As Integer
    Dim tci As Integer
    
    Dim base_verts As Integer
    Dim base_normals As Integer
    Dim base_tex_coords As Integer

    With obj
        .head.NumVerts = .head.NumVerts + 1
        ReDim Preserve .Verts(.head.NumVerts - 1)
        ReDim Preserve .vcolors(.head.NumVerts - 1)
        
        If obj.Groups(Group).texFlag = 1 Then
            .head.NumTexCs = .head.NumTexCs + 1
            ReDim Preserve .TexCoords(.head.NumTexCs - 1)
        End If
        If glIsEnabled(GL_LIGHTING) = GL_TRUE Then
            .head.NumNormals = .head.NumVerts
            ReDim Preserve .Normals(.head.NumNormals - 1)
            .head.NumNormals = .head.NumVerts
            .head.NumNormInds = .head.NumVerts
            ReDim Preserve .NormalIndex(.head.NumNormInds - 1)
            .NormalIndex(.head.NumNormInds - 1) = .head.NumNormInds - 1
        End If
    End With
        
    If Group < obj.head.NumGroups - 1 Then
        With obj
            base_verts = .Groups(Group + 1).offvert
            
            For vi = .head.NumVerts - 1 To base_verts Step -1
                .Verts(vi) = .Verts(vi - 1)
                .vcolors(vi) = .vcolors(vi - 1)
            Next vi
            
            If obj.Groups(Group).texFlag = 1 Then
                base_tex_coords = .Groups(Group).offTex + .Groups(Group).numvert
                
                For tci = .head.NumTexCs - 1 To base_tex_coords Step -1
                    .TexCoords(tci) = .TexCoords(tci - 1)
                Next tci
            End If
            
            If glIsEnabled(GL_LIGHTING) = GL_TRUE Then
                If obj.Groups(Group).texFlag = 1 Then
                    base_normals = .Groups(Group + 1).offvert
                    
                    For ni = .head.NumNormals - 1 To base_normals Step -1
                        .Normals(ni) = .Normals(ni - 1)
                    Next ni
                End If
            End If
        End With
        
        For gi = Group + 1 To obj.head.NumGroups - 1
            With obj.Groups(gi)
                .offvert = .offvert + 1
                If obj.Groups(Group).texFlag = 1 And .texFlag = 1 Then
                    .offTex = .offTex + 1
                End If
            End With
        Next gi
    End If
    
    If Group < obj.head.NumGroups Then
        With obj.Groups(Group)
            obj.Verts(.offvert + .numvert) = v
            obj.vcolors(.offvert + .numvert) = vc
            AddVertex = .offvert + .numvert
            .numvert = .numvert + 1
        End With
    Else
        AddVertex = -1
    End If
End Function
Function AddPolygon(ByRef obj As PModel, ByRef verts_index_buff() As Integer) As Integer
'-------- Warning! Can cause the Normals to be inconsistent if lights are disabled.-----
'---------------------------------Must call ComputeNormals -----------------------------
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    Dim Group As Integer
    
    Dim base_polys As Integer
    
    Dim temp_r As Integer
    Dim temp_g As Integer
    Dim temp_b As Integer
    Dim temp_a As Integer
    
    Dim v_temp As Point3D
    Dim c_temp As color
    
    If verts_index_buff(0) <> verts_index_buff(1) And _
       verts_index_buff(0) <> verts_index_buff(2) Then
        Group = GetVertexGroup(obj.Groups, verts_index_buff(0))
        
        obj.head.NumPolys = obj.head.NumPolys + 1
        ReDim Preserve obj.polys(obj.head.NumPolys - 1)
        ReDim Preserve obj.PColors(obj.head.NumPolys - 1)
        
        If Group < obj.head.NumGroups - 1 Then
            base_polys = obj.Groups(Group + 1).offpoly
            
            For PI = obj.head.NumPolys - 1 To base_polys Step -1
                obj.polys(PI) = obj.polys(PI - 1)
                obj.PColors(PI) = obj.PColors(PI - 1)
            Next PI
            
        
            For gi = Group + 1 To obj.head.NumGroups - 1
                With obj.Groups(gi)
                    .offpoly = .offpoly + 1
                End With
            Next gi
        End If
        
        If Group < obj.head.NumGroups Then
            With obj.polys(obj.Groups(Group).offpoly + obj.Groups(Group).numPoly)
                .Verts(0) = verts_index_buff(0) - obj.Groups(Group).offvert
                If .Verts(0) < 0 Or .Verts(0) >= obj.Groups(Group).numvert Then
                    v_temp = obj.Verts(verts_index_buff(0))
                    c_temp = obj.vcolors(verts_index_buff(0))
                    .Verts(0) = AddVertex(obj, Group, v_temp, c_temp) - obj.Groups(Group).offvert
                End If
                .Verts(1) = verts_index_buff(1) - obj.Groups(Group).offvert
                If .Verts(1) < 0 Or .Verts(1) >= obj.Groups(Group).numvert Then
                    v_temp = obj.Verts(verts_index_buff(1))
                    c_temp = obj.vcolors(verts_index_buff(1))
                    .Verts(1) = AddVertex(obj, Group, v_temp, c_temp) - obj.Groups(Group).offvert
                End If
                .Verts(2) = verts_index_buff(2) - obj.Groups(Group).offvert
                If .Verts(2) < 0 Or .Verts(2) >= obj.Groups(Group).numvert Then
                    v_temp = obj.Verts(verts_index_buff(2))
                    c_temp = obj.vcolors(verts_index_buff(2))
                    .Verts(2) = AddVertex(obj, Group, v_temp, c_temp) - obj.Groups(Group).offvert
                End If
            End With
            
            AddPolygon = obj.Groups(Group).numPoly
            
            
            For vi = 0 To 2
                With obj.vcolors(verts_index_buff(vi))
                    'temp_a = temp_a + .a
                    temp_r = temp_r + .r
                    temp_g = temp_g + .g
                    temp_b = temp_b + .B
                End With
            Next vi
            With obj.Groups(Group)
                obj.PColors(.offpoly + .numPoly).a = 255 'temp_a / 3
                obj.PColors(.offpoly + .numPoly).r = temp_r / 3
                obj.PColors(.offpoly + .numPoly).g = temp_g / 3
                obj.PColors(.offpoly + .numPoly).B = temp_b / 3
                .numPoly = .numPoly + 1
            End With
        Else
            AddPolygon = -1
        End If
    Else
        AddPolygon = -1
    End If
    
    'If Not CheckModelConsistency(obj) Then
    '    'Debug.Print "WTF!!!"
    'End If
End Function
'Removes all polygons refering to the vertex, but doesn't remove the vertex itself
Sub DisableVertex(ByRef obj As PModel, ByVal v_index As Integer)
'------------------------------�WARNINGS!------------------------------
'-------*Causes the Normals to be inconsistent (call ComputeNormals).--
'-------*Causes inconsistent edges (call ComputeEdges).----------------
'-------*Causes unused vertices (call KillUnusedVertices).-------------
    Dim PI As Integer
    
    Dim n_adjacent_polys As Integer
    Dim polys_buff() As Integer
    ReDim polys_buff(obj.head.NumPolys)
    
    n_adjacent_polys = GetAdjacentPolygonsVertex(obj, v_index, polys_buff)
        
    For PI = 0 To n_adjacent_polys - 1
        RemovePolygon obj, polys_buff(PI) - PI
    Next PI
End Sub
Sub RemovePolygon(ByRef obj As PModel, ByVal p_index As Integer)
'------------------------------�WARNINGS!------------------------------
'-------*Causes the Normals to be inconsistent (call ComputeNormals).--
'-------*Causes inconsistent edges (call ComputeEdges).----------------
'-------*Can cause unused vertices (call KillUnusedVertices).----------
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    Dim Group As Integer
    
    Group = GetPolygonGroup(obj.Groups, p_index)
    
    Dim base_verts As Integer
    
    'If obj.head.NumPolys = 1 Then
    '    MsgBox "A P model must have at least 1 polygon. Can't remove this polygon."
    '    Exit Sub
    'End If
    
    With obj
        .head.NumPolys = .head.NumPolys - 1
        
        For PI = p_index To .head.NumPolys - 1
            .polys(PI) = .polys(PI + 1)
            .PColors(PI) = .PColors(PI + 1)
        Next PI
        
        If Group < obj.head.NumGroups - 1 Then
            For gi = Group + 1 To .head.NumGroups - 1
                .Groups(gi).offpoly = .Groups(gi).offpoly - 1
            Next gi
        End If
        .Groups(Group).numPoly = .Groups(Group).numPoly - 1
    End With
    
    'This is technically wrong. The vector shold be emptied if obj.head.NumPolys droped to 0,
    'but they should be inmediately refilled with something else because a P Model can't have 0
    'polygons.
    If obj.head.NumPolys >= 1 Then
        ReDim Preserve obj.polys(obj.head.NumPolys - 1)
        ReDim Preserve obj.PColors(obj.head.NumPolys - 1)
    End If
End Sub
Sub RemoveGroup(ByRef obj As PModel, ByVal g_index As Integer)
    Dim gi As Integer
    Dim offvert As Integer
    Dim offpoly As Integer
    Dim offEdge As Integer
    Dim offTex As Integer
    
    If (obj.Groups(g_index).numvert > 0) Then
        RemoveGroupVColors obj, g_index
        RemoveGroupVertices obj, g_index
        RemoveGroupPColors obj, g_index
        RemoveGroupPolys obj, g_index
        RemoveGroupEdges obj, g_index
        RemoveGroupTexCoords obj, g_index
    Else
        ''Debug.Print obj.Groups(g_index).numvert; " "; obj.Groups(g_index).numPoly
    End If
    
    RemoveGroupHundret obj, g_index
    RemoveGroupHeader obj, g_index
    
    For gi = g_index To obj.head.NumGroups - 1
        obj.Groups(gi) = obj.Groups(gi + 1)
        If gi > 0 Then
            With obj.Groups(gi - 1)
                obj.Groups(gi).offvert = .offvert + .numvert
                obj.Groups(gi).offpoly = .offpoly + .numPoly
                obj.Groups(gi).offEdge = .offEdge + .numEdge
                If .texFlag = 1 Then
                    obj.Groups(gi).offTex = .offTex + .numvert
                Else
                    obj.Groups(gi).offTex = .offTex
                End If
            End With
        Else
            With obj.Groups(gi)
                .offvert = 0
                .offpoly = 0
                .offTex = 0
                .offEdge = 0
            End With
        End If
    Next gi
    
    ReDim Preserve obj.Groups(obj.head.NumGroups - 1)
    
    'If Not CheckModelConsistency(obj) Then
    '    'Debug.Print "WTF!!!"
    'End If
    
    ComputeNormals obj
End Sub
Sub RemoveGroupVertices(ByRef obj As PModel, g_index As Integer)
    Dim vi As Integer
    Dim vi2 As Integer
    
    With obj
        If g_index < obj.head.NumGroups - 1 Then
            vi2 = .Groups(g_index).offvert
            For vi = obj.Groups(g_index + 1).offvert To obj.head.NumVerts - 1
                obj.Verts(vi2) = obj.Verts(vi)
                vi2 = vi2 + 1
            Next vi
        End If
        
        ReDim Preserve .Verts(.head.NumVerts - .Groups(g_index).numvert)
    End With
End Sub
Sub RemoveGroupVColors(ByRef obj As PModel, g_index As Integer)
    Dim vci As Integer
    Dim vci2 As Integer
    
    With obj
        If g_index < obj.head.NumGroups - 1 Then
            vci2 = .Groups(g_index).offvert
            For vci = obj.Groups(g_index + 1).offvert To obj.head.NumVerts - 1
                obj.vcolors(vci2) = obj.vcolors(vci)
                vci2 = vci2 + 1
            Next vci
        End If
        
        ReDim Preserve .vcolors(.head.NumVerts - .Groups(g_index).numvert)
    End With
End Sub
Sub RemoveGroupPolys(ByRef obj As PModel, g_index As Integer)
    Dim PI As Integer
    Dim pi2 As Integer
    
    With obj
        If g_index < .head.NumGroups - 1 Then
            pi2 = .Groups(g_index).offpoly
            For PI = .Groups(g_index + 1).offpoly To .head.NumPolys - 1
                .polys(pi2) = .polys(PI)
                pi2 = pi2 + 1
            Next PI
        End If
        
        ReDim Preserve .polys(.head.NumPolys - .Groups(g_index).numPoly)
    End With
End Sub
Sub RemoveGroupPColors(ByRef obj As PModel, g_index As Integer)
    Dim pci As Integer
    Dim pci2 As Integer
    
    With obj
        If g_index < obj.head.NumGroups - 1 Then
            pci2 = .Groups(g_index).offpoly
            For pci = obj.Groups(g_index + 1).offpoly To obj.head.NumPolys - 1
                obj.PColors(pci2) = obj.PColors(pci)
                pci2 = pci2 + 1
            Next pci
        End If
        
        ReDim Preserve .PColors(.head.NumPolys - .Groups(g_index).numPoly)
    End With
End Sub
Sub RemoveGroupEdges(ByRef obj As PModel, g_index As Integer)
    Dim ei As Integer
    Dim ei2 As Integer
    
    With obj
        If g_index < obj.head.NumGroups - 1 Then
            ei2 = .Groups(g_index).offEdge
            For ei = obj.Groups(g_index + 1).offEdge To obj.head.NumEdges - 1
                obj.Edges(ei2) = obj.Edges(ei)
                ei2 = ei2 + 1
            Next ei
        End If
        
        ReDim Preserve .Edges(.head.NumEdges - .Groups(g_index).numEdge)
    End With
End Sub
Sub RemoveGroupTexCoords(ByRef obj As PModel, g_index As Integer)
    Dim ti As Integer
    Dim ti2 As Integer
    
    With obj
        If .Groups(g_index).texFlag = 1 Then
            'If .Groups(g_index).polyType > 1 Then
                If g_index < obj.head.NumGroups - 1 Then
                    ti2 = .Groups(g_index).offTex
                    For ti = obj.Groups(g_index + 1).offTex To obj.head.NumTexCs - 1
                        obj.TexCoords(ti2) = obj.TexCoords(ti)
                        ti2 = ti2 + 1
                    Next ti
                End If
                
                If obj.Groups(g_index).texFlag = 1 Then _
                    ReDim Preserve .TexCoords(.head.NumTexCs - .Groups(g_index).numvert)
            'End If
        End If
    End With
End Sub
Sub RemoveGroupHundret(ByRef obj As PModel, g_index As Integer)
    Dim hi As Integer
    Dim hi2 As Integer
    
    With obj
        If g_index < obj.head.NumGroups - 1 Then
            For hi = g_index + 1 To obj.head.NumGroups - 1
                obj.hundrets(hi - 1) = obj.hundrets(hi)
            Next hi
        End If
        
        ReDim Preserve .hundrets(.head.mirex_h - 2)
    End With
End Sub
Sub RemoveGroupHeader(ByRef obj As PModel, g_index As Integer)
    With obj.head
        .NumPolys = .NumPolys - obj.Groups(g_index).numPoly
        .NumEdges = .NumEdges - obj.Groups(g_index).numEdge
        .NumVerts = .NumVerts - obj.Groups(g_index).numvert
        .mirex_h = .mirex_h - 1
        If obj.Groups(g_index).texFlag = 1 Then _
            .NumTexCs = .NumTexCs - obj.Groups(g_index).numvert
        .NumGroups = .NumGroups - 1
    End With
End Sub

Sub PaintPolygon(ByRef obj As PModel, ByVal p_index As Integer, ByVal r As Byte, ByVal g As Byte, ByVal B As Byte)
'------------------------------�WARNINGS!----------------------------------
'-------*Can causes the Normals to be inconsistent (call ComputeNormals).--
'-------*Can causes inconsistent edges (call ComputeEdges).----------------
'-------*Can cause unused vertices (call KillUnusedVertices).--------------
    Dim Group As Integer
    Dim vi As Integer
    
    Group = GetPolygonGroup(obj.Groups, p_index)
    For vi = 0 To 2
        With obj.polys(p_index)
            .Verts(vi) = PaintVertex(obj, Group, .Verts(vi), r, g, B, obj.Groups(Group).texFlag <> 0)
            ''Debug.Print "Vert(:", .Verts(vi), ",", Group, ")", obj.Verts(.Verts(vi) + obj.Groups(Group).offVert).x, obj.Verts(.Verts(vi) + obj.Groups(Group).offVert).y, obj.Verts(.Verts(vi) + obj.Groups(Group).offVert).z
        End With
        With obj.PColors(p_index)
            .r = r
            .g = g
            .B = B
        End With
    Next vi
End Sub
Function PaintVertex(ByRef obj As PModel, ByVal g_index As Integer, ByVal v_index As Integer, ByVal r As Byte, ByVal g As Byte, ByVal B As Byte, ByVal Textured As Boolean) As Integer
    Dim vi As Integer
    Dim n_verts As Integer
    Dim v_list() As Integer
    
    Dim v_temp As Point3D
    Dim c_temp As color
    
    Dim tc_tmp As Point2D
    
    PaintVertex = -1
    
    With obj.vcolors(v_index + obj.Groups(g_index).offvert)
        If .r = r And .g = g And .B = B Then
            PaintVertex = v_index
        End If
    End With
    
    If PaintVertex = -1 Then
        n_verts = GetEqualGroupVertices(obj, v_index + obj.Groups(g_index).offvert, v_list)
        
        For vi = 0 To n_verts - 1
            ''Debug.Print "Found("; vi; ")"; v_list(vi)
            With obj.vcolors(v_list(vi))
                If .r = r And .g = g And .B = B Then
                    PaintVertex = v_list(vi) - obj.Groups(g_index).offvert
                    Exit For
                End If
            End With
        Next vi
        
        If PaintVertex = -1 Then
            With obj.Verts(v_index + obj.Groups(g_index).offvert)
                v_temp.x = .x
                v_temp.y = .y
                v_temp.z = .z
            End With
            
            With c_temp
                .r = r
                .g = g
                .B = B
            End With
            
            If Textured Then _
                tc_tmp = obj.TexCoords(obj.Groups(g_index).offTex + v_index)
                
            PaintVertex = AddVertex(obj, g_index, v_temp, c_temp) - obj.Groups(g_index).offvert
            
            If glIsEnabled(GL_LIGHTING) = GL_TRUE Then _
                obj.Normals(PaintVertex) = obj.Normals(obj.Groups(g_index).offvert + v_index)
            
            If Textured Then _
                obj.TexCoords(obj.Groups(g_index).offTex + PaintVertex) = tc_tmp
            
            'obj.Normals(PaintVertex + obj.Groups(g_index).offvert) = obj.Normals(v_index)
            
        Else
            ''Debug.Print "Substituido por: " + Str$(PaintVertex)
        End If
    End If
End Function
Sub ResizeModel(ByRef obj As PModel, ByVal redX As Double, ByVal redY As Double, ByVal redZ As Double)
    Dim vi
    
    For vi = 0 To obj.head.NumVerts - 1
        With obj.Verts(vi)
            .x = .x * redX
            .y = .y * redY
            .z = .z * redZ
        End With
    Next vi
End Sub
Sub ApplyCurrentVColors(ByRef obj As PModel)
    Dim gi As Integer
    Dim vi As Integer
    Dim vp(4) As Long
    
    glDisable GL_BLEND

    For gi = 0 To obj.head.NumGroups - 1
        With obj.Groups(gi)
            For vi = .offvert To .offvert + .numvert - 1
                obj.vcolors(vi) = GetVertColor(obj.Verts(vi), obj.Normals(vi), obj.vcolors(vi))
            Next vi
        End With
    Next gi
End Sub
Sub ApplyCurrentVCoords(ByRef obj As PModel)
    Dim vi As Integer
    
    For vi = 0 To obj.head.NumVerts - 1
        obj.Verts(vi) = GetEyeSpaceCoords(obj.Verts(vi))
    Next vi
End Sub
'----------------------------------------------------------------------------------------------------
'=============================================TOPOLOGIC==============================================
'----------------------------------------------------------------------------------------------------
Function GetAdjacentPolygonsVertex(ByRef obj As PModel, ByVal v_index As Integer, ByRef poly_buff() As Integer)
    Dim Group As Integer
    Dim PI As Integer
    Dim vi As Integer
    
    Dim n_polys As Integer

    With obj
        Group = GetVertexGroup(obj.Groups, v_index)
    
        n_polys = 0
        For PI = .Groups(Group).offpoly To .Groups(Group).offpoly + .Groups(Group).numPoly - 1
            For vi = 0 To 2
                If .polys(PI).Verts(vi) = v_index - .Groups(Group).offvert Then
                    ReDim Preserve poly_buff(n_polys)
                    poly_buff(n_polys) = PI
                    n_polys = n_polys + 1
                    Exit For
                End If
            Next vi
        Next PI
    End With
        
    GetAdjacentPolygonsVertex = n_polys
End Function
Function GetAdjacentPolygonsVertices(ByRef obj As PModel, ByRef v_indices() As Integer, ByRef poly_buff() As Integer)
    Dim Group As Integer
    Dim PI As Integer
    Dim vi As Integer
    Dim pvi As Integer
    Dim n_verts As Integer
    Dim n_polys As Integer
    
    n_verts = UBound(v_indices) + 1

    n_polys = 0
    For vi = 0 To n_verts - 1
        Group = GetVertexGroup(obj.Groups, v_indices(vi))
        With obj.Groups(Group)
            For PI = .offpoly To .offpoly + .numPoly - 1
                For pvi = 0 To 2
                    If obj.polys(PI).Verts(pvi) = v_indices(vi) - .offvert Then
                        ReDim Preserve poly_buff(n_polys)
                        poly_buff(n_polys) = PI
                        n_polys = n_polys + 1
                        Exit For
                    End If
                Next pvi
            Next PI
        End With
    Next vi
        
    GetAdjacentPolygonsVertices = n_polys
End Function
Function GetPolygonAdjacentVertexIndices(obj As PModel, ByRef v_polygons() As Integer, ByRef v_indices_discarded() As Integer, ByRef v_adjacent_indices_out() As int_vector) As Integer
    Dim Group As Integer
    Dim PI As Integer
    Dim pvi As Integer
    Dim vid As Integer
    Dim num_polys As Integer
    Dim num_verts As Integer
    Dim num_discardeds As Integer
    Dim num_equal_verts As Integer
    Dim off_vert As Integer
    Dim equal_verts() As Integer
    Dim foundQ As Boolean

    num_discardeds = UBound(v_indices_discarded) + 1
    num_polys = UBound(v_polygons) + 1
    
    GetPolygonAdjacentVertexIndices = 0
    
    For PI = 0 To num_polys - 1
        Group = GetPolygonGroup(obj.Groups, v_polygons(PI))
        off_vert = obj.Groups(Group).offvert
        With obj.polys(v_polygons(PI))
            For pvi = 0 To 2
                'Check whether the vertex should be ignored or not
                foundQ = False
                For vid = 0 To num_discardeds - 1
                    If v_indices_discarded(vid) - off_vert = .Verts(pvi) Then
                        foundQ = True
                        Exit For
                    End If
                Next vid
                If Not foundQ Then
                    'Check if the vertex (or similar) is already added to the list

                    For vid = 0 To GetPolygonAdjacentVertexIndices - 1
                        If ComparePoints3D(obj.Verts(v_adjacent_indices_out(vid).vector(0)), _
                                            obj.Verts(.Verts(pvi) + off_vert)) Then
                            foundQ = True
                            Exit For
                        End If
                    Next vid
                        
                    If Not foundQ Then
                        ReDim equal_verts(0)
                        'Find all similar vertices
                        num_equal_verts = GetEqualVertices(obj, .Verts(pvi), equal_verts)
                        'Update the output data
                        GetPolygonAdjacentVertexIndices = GetPolygonAdjacentVertexIndices + 1
                        ReDim Preserve v_adjacent_indices_out(GetPolygonAdjacentVertexIndices - 1)
                        ReDim v_adjacent_indices_out(GetPolygonAdjacentVertexIndices - 1).vector(num_equal_verts - 1)
                        v_adjacent_indices_out(GetPolygonAdjacentVertexIndices - 1).length = num_equal_verts
                        CopyMemory v_adjacent_indices_out(GetPolygonAdjacentVertexIndices - 1).vector(0) _
                                    , equal_verts(0), num_equal_verts * 2
                        'obj.vcolors(v_adjacent_indices_out(GetPolygonAdjacentVertexIndices - 1).vector(0)).r = 255
                        'obj.vcolors(v_adjacent_indices_out(GetPolygonAdjacentVertexIndices - 1).vector(0)).g = 0
                        'obj.vcolors(v_adjacent_indices_out(GetPolygonAdjacentVertexIndices - 1).vector(0)).b = 0
                    End If
                End If
            Next pvi
        End With
    Next PI

End Function
Function GetEqualGroupVertices(ByRef obj As PModel, ByVal v_index As Integer, ByRef v_list() As Integer) As Integer
    Dim vi As Integer
    Dim Group As Integer
    
    Dim n_verts As Integer
    Dim v As Point3D
    
    v = obj.Verts(v_index)
    Group = GetVertexGroup(obj.Groups, v_index)
    For vi = obj.Groups(Group).offvert To obj.Groups(Group).offvert + obj.Groups(Group).numvert - 1
        If ComparePoints3D(obj.Verts(vi), v) Then
            ReDim Preserve v_list(n_verts)
            v_list(n_verts) = vi
            ''Debug.Print "Intended("; n_verts; ")"; Str$(vi)
            n_verts = n_verts + 1
        End If
    Next vi
    
    GetEqualGroupVertices = n_verts
End Function
Function GetEqualVertices(ByRef obj As PModel, ByVal v_index As Integer, ByRef v_list() As Integer) As Integer
    Dim vi As Integer
    Dim gi As Integer
    
    Dim n_verts As Integer
    Dim v As Point3D
    
    v = obj.Verts(v_index)
    For gi = 0 To obj.head.NumGroups - 1
        If Not obj.Groups(gi).HiddenQ Then
            For vi = 0 To obj.head.NumVerts - 1
                If ComparePoints3D(obj.Verts(vi), v) Then
                    ReDim Preserve v_list(n_verts)
                    v_list(n_verts) = vi
                    n_verts = n_verts + 1
                End If
            Next vi
        End If
    Next gi
    
    GetEqualVertices = n_verts
End Function
Public Sub ApplyPChanges(ByRef obj As PModel, ByVal DNormals As Boolean)
    On Error GoTo hand
    
1:    KillUnusedVertices obj
2:    ApplyCurrentVCoords obj
3:    ComputePColors obj
4:    ComputeEdges obj
    
    If DNormals Then
5:        DisableNormals obj
    Else
6:        ComputeNormals obj
    End If
    
    ComputeBoundingBox obj
    
    With obj
        .ResizeX = 1
        .ResizeY = 1
        .ResizeZ = 1
        .RepositionX = 0
        .RepositionY = 0
        .RepositionZ = 0
        .RotateAlpha = 0
        .RotateBeta = 0
        .RotateGamma = 0
        .RotationQuaternion.x = 0
        .RotationQuaternion.y = 0
        .RotationQuaternion.z = 0
        .RotationQuaternion.w = 1
    End With
    Exit Sub
hand:
    MsgBox "PChange at " + obj.fileName + "!!!" + Str$(Erl), vbOKOnly, "Error PChange"
End Sub
Public Sub MoveVertex(ByRef obj As PModel, ByVal v_index As Integer, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Dim p_temp As Point3D
    
    With p_temp
        .x = x
        .y = y
        .z = z
    End With
    
    p_temp = GetUnProjectedCoords(p_temp)

    With obj.Verts(v_index)
        .x = p_temp.x
        .y = p_temp.y
        .z = p_temp.z
    End With
End Sub
Public Sub GetAllNormalDependentPolys(ByRef obj As PModel, ByRef v_indices() As Integer, ByRef adjacent_polys_indices() As Integer, _
                                                        ByRef v_adjacent_verts_indices() As int_vector, ByRef adjacent_adjacent_polys_indices() As int_vector)
    Dim vi As Integer
    Dim pvi As Integer
    Dim num_verts As Integer
    Dim num_polys As Integer
    Dim temp_num_polys As Integer
    Dim temp_polys_indices() As Integer
    Dim n_adjacent_verts As Integer
    
    'Get the polygons adjacent to the selected vertices
    GetAdjacentPolygonsVertices obj, v_indices, adjacent_polys_indices
    
    'Get the vertices adjacent to the selected vertices
    n_adjacent_verts = GetPolygonAdjacentVertexIndices(obj, adjacent_polys_indices, v_indices, v_adjacent_verts_indices)
    
    'Get polygons adjacent to the adjacent
    ReDim adjacent_adjacent_polys_indices(n_adjacent_verts - 1)
    For vi = 0 To n_adjacent_verts - 1
        num_polys = GetAdjacentPolygonsVertices(obj, v_adjacent_verts_indices(vi).vector, adjacent_adjacent_polys_indices(vi).vector)
        adjacent_adjacent_polys_indices(vi).length = num_polys
    Next vi
End Sub
Public Sub UpdateNormals(ByRef obj As PModel, ByRef v_indices() As Integer, ByRef adjacent_polys_indices() As Integer, _
                                                        ByRef v_adjacent_verts_indices() As int_vector, ByRef adjacent_adjacent_polys_indices() As int_vector)
    Dim vi As Integer
    Dim num_adjacents As Integer
    
    UpdateNormal obj, v_indices, adjacent_polys_indices
    
    num_adjacents = UBound(v_adjacent_verts_indices) + 1
    For vi = 0 To num_adjacents - 1
        UpdateNormal obj, v_adjacent_verts_indices(vi).vector, adjacent_adjacent_polys_indices(vi).vector
    Next vi
End Sub
Public Sub UpdateNormal(ByRef obj As PModel, ByRef v_indices() As Integer, ByRef adjacent_polys_indices() As Integer)
    Dim PI As Integer
    Dim vi As Integer
    Dim num_polys As Integer
    Dim num_verts As Integer
    Dim Group As Integer
    
    Dim CurrentNormal As Point3D
    Dim TotalNormal As Point3D
    
    Dim offvert As Integer
    
    With TotalNormal
        .x = 0
        .y = 0
        .z = 0
    End With

    num_polys = UBound(adjacent_polys_indices) + 1
    num_verts = UBound(v_indices) + 1
    
    For PI = 0 To num_polys - 1
        Group = GetPolygonGroup(obj.Groups, adjacent_polys_indices(PI))
        offvert = obj.Groups(Group).offvert
            
        With obj.polys(adjacent_polys_indices(PI))
            CurrentNormal = CalculateNormal(obj.Verts(.Verts(2) + offvert), _
                                            obj.Verts(.Verts(1) + offvert), _
                                            obj.Verts(.Verts(0) + offvert))
        End With
        With TotalNormal
            .x = .x + CurrentNormal.x
            .y = .y + CurrentNormal.y
            .z = .z + CurrentNormal.z
        End With
    Next PI
    
    With TotalNormal
        .x = .x / num_polys
        .y = .y / num_polys
        .z = .z / num_polys
    End With
    TotalNormal = Normalize(TotalNormal)
    
    For vi = 0 To num_verts - 1
        obj.Normals(v_indices(vi)) = TotalNormal
    Next vi
End Sub
Public Function CheckModelConsistency(ByRef obj As PModel) As Boolean
    Dim num_gropus As Integer
    Dim num_polys As Integer
    Dim num_verts As Integer
    Dim num_textures As Integer
    Dim num_norm_inds As Integer
    Dim num_normals As Integer
    Dim off_poly As Integer
    Dim off_vert As Integer
    Dim off_tex As Integer
    Dim end_group_polys As Integer
    Dim end_group_verts As Integer
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    
    CheckModelConsistency = True
    With obj
        num_norm_inds = UBound(.NormalIndex)
        num_normals = UBound(.Normals)
        num_textures = UBound(.TexCoords)
        For gi = 0 To .head.NumGroups - 1
            off_vert = .Groups(gi).offvert
            end_group_verts = .Groups(gi).numvert - 1
            off_poly = .Groups(gi).offpoly
            end_group_polys = .Groups(gi).offpoly + .Groups(gi).numPoly - 1
            off_tex = .Groups(gi).offTex
            For PI = off_poly To end_group_polys
                If (.polys(PI).Verts(0) < 0 Or .polys(PI).Verts(0) > end_group_verts) Then
                    'Debug.Print "ERROR!!!!"
                    CheckModelConsistency = False
                End If
                If (.polys(PI).Verts(1) < 0 Or .polys(PI).Verts(1) > end_group_verts) Then
                    'Debug.Print "ERROR!!!!"
                    CheckModelConsistency = False
                End If
                If (.polys(PI).Verts(2) < 0 Or .polys(PI).Verts(2) > end_group_verts) Then
                    'Debug.Print "ERROR!!!!"
                    CheckModelConsistency = False
                End If
                
                If (.polys(PI).Normals(0) > num_norm_inds) Then
                    'Debug.Print "ERROR!!!!"
                    CheckModelConsistency = False
                ElseIf (num_normals > 0 And .NormalIndex(.polys(PI).Normals(0)) > num_normals) Then
                    'Debug.Print "ERROR!!!!"
                    CheckModelConsistency = False
                End If
                If (.polys(PI).Normals(1) > num_norm_inds) Then
                    'Debug.Print "ERROR!!!!"
                    CheckModelConsistency = False
                ElseIf (num_normals > 0 And .NormalIndex(.polys(PI).Normals(1)) > num_normals) Then
                    'Debug.Print "ERROR!!!!"
                    CheckModelConsistency = False
                End If
                If (.polys(PI).Normals(2) > num_norm_inds) Then
                    'Debug.Print "ERROR!!!!"
                    CheckModelConsistency = False
                ElseIf (num_normals > 0 And .NormalIndex(.polys(PI).Normals(2)) > num_normals) Then
                    'Debug.Print "ERROR!!!!"
                    CheckModelConsistency = False
                End If
                
                If (.Groups(gi).texFlag = 1) Then
                    If (.polys(PI).Verts(0) + off_tex > num_textures) Then
                        'Debug.Print "ERROR!!!!"
                        CheckModelConsistency = False
                    End If
                    If (.polys(PI).Verts(1) + off_tex > num_textures) Then
                        'Debug.Print "ERROR!!!!"
                        CheckModelConsistency = False
                    End If
                    If (.polys(PI).Verts(2) + off_tex > num_textures) Then
                        'Debug.Print "ERROR!!!!"
                        CheckModelConsistency = False
                    End If
                End If
            Next PI
        Next gi
    End With
End Function
Function CutPolygonThroughPlane(ByRef obj As PModel, ByVal p_index As Integer, _
                                ByVal g_index As Integer, ByVal a As Single, ByVal B As Single, _
                                ByVal C As Single, ByVal d As Single, _
                                ByRef known_plane_pointsV() As Point3D) As Boolean
    
    Dim vi As Integer
    Dim PI As Integer
    Dim num_known_plane_points As Integer
    
    Dim offTex As Integer
    Dim offvert As Long
    Dim polyNormal As Point3D
    Dim isParalelQ As Boolean
    Dim equality As Single
    Dim equalityValidQ As Boolean
    Dim d_poly As Single
    Dim ei As Integer
    Dim lambda_mult_plane As Single
    Dim k_plane As Single
    Dim alpha_plane As Single
    Dim lambda_mult_poly As Single
    Dim k_poly As Single
    Dim alpha_poly As Single
    Dim cutQ As Boolean
    Dim p1_index As Integer
    Dim p2_index As Integer
    Dim t1_index As Integer
    Dim t2_index As Integer
    Dim p1 As Point3D
    Dim p2 As Point3D
    
    Dim g_index_old As Integer
    
    Dim cut_point As Point3D
    
    Dim v_index_rectify As Long
    Dim num_equal_verts As Integer
    Dim v_equal_indicesV() As Integer
    Dim p_rect As Point3D
    
    
    Dim intersection_point As Point3D
    Dim intersection_tex_coord As Point2D
    
    Dim p1IsContainedQ As Boolean
    Dim p2IsContainedQ As Boolean
    
    g_index_old = g_index

    offvert = obj.Groups(g_index).offvert
    offTex = obj.Groups(g_index).offTex
    With obj.polys(p_index)
        polyNormal = CalculateNormal(obj.Verts(.Verts(0) + offvert), obj.Verts(.Verts(1) + offvert), obj.Verts(.Verts(2) + offvert))
        polyNormal = Normalize(polyNormal)
    End With
    
    If SafeArrayGetDim(known_plane_pointsV) <> 0 Then
        num_known_plane_points = UBound(known_plane_pointsV) + 1
    Else
        num_known_plane_points = 0
    End If
    
    'Check wether the planes are paralel or not.
    'If they are, don't cut the polygon.
    isParalelQ = True
    equalityValidQ = False
    If (polyNormal.x = 0 Or a = 0) Then
        isParalelQ = Abs(polyNormal.x - a) < 0.0001
    Else
        equalityValidQ = True
        equality = a / polyNormal.x
    End If
    
    If (polyNormal.y = 0 Or B = 0) Then
        isParalelQ = isParalelQ And Abs(polyNormal.y - B) < 0.0001
    Else
        If (equalityValidQ) Then
            isParalelQ = isParalelQ And Abs((B / polyNormal.y) - equality) < 0.0001
        Else
            equalityValidQ = True
            equality = B / polyNormal.y
        End If
    End If
    
    If (polyNormal.z = 0 Or C = 0) Then
        isParalelQ = isParalelQ And (polyNormal.z = C)
    Else
        If (equalityValidQ) Then
            isParalelQ = isParalelQ And Abs((C / polyNormal.z) - equality) < 0.0001
        Else
            equalityValidQ = True
            equality = C / polyNormal.z
        End If
    End If
    
    If Not isParalelQ Then
        With obj.Verts(obj.polys(p_index).Verts(0) + offvert)
            d_poly = -polyNormal.x * .x - polyNormal.y * .y - polyNormal.z * .z
        End With
        
        ei = 0
        cutQ = False
        Do
            p1_index = obj.polys(p_index).Verts(ei) + offvert
            p2_index = obj.polys(p_index).Verts((ei + 1) Mod 3) + offvert
            
            t1_index = obj.polys(p_index).Verts(ei) + offTex
            t2_index = obj.polys(p_index).Verts((ei + 1) Mod 3) + offTex
            If ComparePoints3D(obj.Verts(p2_index), obj.Verts(p1_index)) Then
                'Degenerated triangle, don't bother
                CutPolygonThroughPlane = False
                Exit Function
            End If
                
            'Check if the edge is contained on the plane
            p1IsContainedQ = False
            p2IsContainedQ = False
            
            For PI = 0 To num_known_plane_points - 1
                If ComparePoints3D(obj.Verts(p1_index), known_plane_pointsV(PI)) Then
                    p1IsContainedQ = True
                End If
                If ComparePoints3D(obj.Verts(p2_index), known_plane_pointsV(PI)) Then
                    p2IsContainedQ = True
                End If
                
                If (p1IsContainedQ And p2IsContainedQ) Then
                    CutPolygonThroughPlane = False
                    Exit Function
                End If
            Next PI
            
            'If they aren't, find the cut point.
            With obj.Verts(p1_index)
                lambda_mult_plane = -a * .x - B * .y - C * .z
                k_plane = lambda_mult_plane - d
            End With
            
            With obj.Verts(p2_index)
                lambda_mult_plane = lambda_mult_plane + a * .x + B * .y + C * .z
            End With
            If (Abs(lambda_mult_plane) > 0.0000001 And k_plane <> 0) Then
                alpha_plane = k_plane / lambda_mult_plane
                intersection_point = CalculateLinePoint(alpha_plane, _
                                        obj.Verts(p1_index), obj.Verts(p2_index))
                
                If (obj.Groups(g_index).texFlag = 1) Then _
                    intersection_tex_coord = GetPointInLine2D(obj.TexCoords(t1_index), _
                                                obj.TexCoords(t2_index), alpha_plane)
                
                'Finally check if cut point is actually inside the edge segment.
                If (alpha_plane > 0.2 And alpha_plane < 0.8) Then
                    cutQ = CutEdgeAtPoint(obj, p_index, ei, intersection_point, intersection_tex_coord)
                    CheckModelConsistency obj
                    g_index = GetPolygonGroup(obj.Groups, p_index)
                    While FindNextAdjacentPolyEdgeForward(EditedPModel, obj.Verts(p1_index), _
                                                   obj.Verts(p2_index), g_index, p_index, ei)
                        'Must recompute the texture junction point everytime we go beyond a textured
                        'group boundaries.
                        If g_index_old <> g_index Then
                            If (obj.Groups(g_index).texFlag = 1) Then
                                offTex = obj.Groups(g_index).offTex
                                t1_index = obj.polys(p_index).Verts(ei) + offTex
                                t2_index = obj.polys(p_index).Verts((ei + 1) Mod 3) + offTex
                                intersection_tex_coord = GetPointInLine2D(obj.TexCoords(t1_index), _
                                                obj.TexCoords(t2_index), alpha_plane)
                            End If
                            g_index_old = g_index
                        End If
                        cutQ = CutEdgeAtPoint(EditedPModel, p_index, ei, intersection_point, intersection_tex_coord)
                    Wend
                    'Add the new point to the known plane points list
                    ReDim known_plane_pointsV(num_known_plane_points)
                    known_plane_pointsV(num_known_plane_points) = intersection_point
                    'Just one cut per polygon. After cutting an edge, exit the loop.
                Else
                    'If it's close enough, change the vertex location so that it's contained on the plane
                    If alpha_plane <= 0.2 And alpha_plane >= 0 Then
                        v_index_rectify = p1_index
                    ElseIf alpha_plane >= 0.8 And alpha_plane <= 1 Then
                        v_index_rectify = p2_index
                    Else
                        v_index_rectify = -1
                    End If
                    
                    If v_index_rectify <> -1 Then
                        'Add the rectified point to the known plane points list
                        ReDim known_plane_pointsV(num_known_plane_points)
                        known_plane_pointsV(num_known_plane_points) = intersection_point
                        num_known_plane_points = num_known_plane_points + 1
                    
                        num_equal_verts = GetEqualVertices(obj, v_index_rectify, v_equal_indicesV)
                            
                        'Propagate changes to all equal vertices
                        For vi = 0 To num_equal_verts - 1
                            obj.Verts(v_equal_indicesV(vi)) = intersection_point
                        Next vi
                        'cutQ = True
                        'Exit Do
                    End If
                End If
            End If
            ei = ei + 1
        Loop Until cutQ Or ei > 2
    End If
    
    CutPolygonThroughPlane = cutQ
End Function

Sub CutPModelThroughPlane(ByRef obj As PModel, ByVal a As Single, ByVal B As Single, _
                            ByVal C As Single, ByVal d As Single, _
                            ByRef known_plane_pointsV() As Point3D)
    Dim gi As Integer
    Dim PI As Integer
    Dim vi As Integer
    
    Dim offpoly As Long
    Dim offvert As Long
    
    For gi = 0 To obj.head.NumGroups - 1
        offpoly = obj.Groups(gi).offpoly
        PI = offpoly
        While PI < offpoly + obj.Groups(gi).numPoly - 1
            If Not CutPolygonThroughPlane(obj, PI, gi, a, B, C, d, known_plane_pointsV) Then
                CheckModelConsistency obj
                PI = PI + 1
            End If
        Wend
    Next gi

End Sub

Sub EraseEmisphereVertices(ByRef obj As PModel, ByVal a As Single, ByVal B As Single, _
                            ByVal C As Single, ByVal d As Single, _
                            ByVal underPlaneQ As Boolean, _
                            ByRef known_plane_pointsV() As Point3D)

    Dim gi As Integer
    Dim PI As Long
    Dim vi As Long
    Dim offvert As Long
    Dim offpoly As Long
    Dim atLeastOneSparedQ As Boolean
    
    Dim v_index As Integer
    Dim kppi As Integer
    Dim num_known_plane_points As Integer
    Dim foundQ As Boolean
    
    If SafeArrayGetDim(known_plane_pointsV) <> 0 Then
        num_known_plane_points = UBound(known_plane_pointsV) + 1
    Else
        num_known_plane_points = 0
    End If
    
    For gi = 0 To obj.head.NumGroups - 1
        offvert = obj.Groups(gi).offvert
        offpoly = obj.Groups(gi).offpoly
        
        PI = offpoly
        While PI < offpoly + obj.Groups(gi).numPoly And obj.head.NumPolys > 1
            atLeastOneSparedQ = False
            For vi = 0 To 2
                foundQ = False
                v_index = obj.polys(PI).Verts(vi) + offvert
                For kppi = 0 To num_known_plane_points - 1
                    If ComparePoints3D(obj.Verts(v_index), known_plane_pointsV(kppi)) Then
                        foundQ = True
                        Exit For
                    End If
                Next kppi
                If Not foundQ Then
                    If underPlaneQ Then
                        If IsPoint3DUnderPlane(obj.Verts(v_index), a, B, C, d) Then
                            atLeastOneSparedQ = True
                        End If
                    Else
                        If IsPoint3DAbovePlane(obj.Verts(v_index), a, B, C, d) Then
                            atLeastOneSparedQ = True
                        End If
                    End If
                End If
            Next vi
            If Not atLeastOneSparedQ Then
                RemovePolygon obj, PI
            Else
                PI = PI + 1
            End If
        Wend
    Next gi
    
    If obj.head.NumPolys = 1 Then _
        MsgBox "A P model must have at least one polygon. The last triangle was spared."
    
    KillUnusedVertices obj
    KillEmptyGroups obj
End Sub

Sub MirrorEmisphere(ByRef obj As PModel, ByVal a As Single, ByVal B As Single, ByVal C As Single, _
                    ByVal d As Single)
    Dim gi As Integer
    Dim num_groups As Integer
    
    num_groups = obj.head.NumGroups
    
    For gi = 0 To num_groups - 1
        If Not obj.Groups(gi).HiddenQ Then _
            MirrorGroupRelativeToPlane obj, gi, a, B, C, d
    Next gi
End Sub

Sub DuplicateMirrorEmisphere(ByRef obj As PModel, ByVal a As Single, ByVal B As Single, _
                            ByVal C As Single, ByVal d As Single)
    Dim gi As Integer
    Dim gi_mirror As Integer
    Dim num_groups As Integer
    
    num_groups = obj.head.NumGroups
    
    gi_mirror = num_groups
    For gi = 0 To num_groups - 1
        If DuplicateGroup(obj, gi) Then
            MirrorGroupRelativeToPlane obj, gi_mirror, a, B, C, d
            gi_mirror = gi_mirror + 1
        End If
    Next gi
End Sub

Function DuplicateGroup(ByRef obj As PModel, ByVal g_index As Integer) As Boolean
    Dim vertsV() As Point3D
    Dim facesV() As PPolygon
    Dim TexCoordsV() As Point2D
    Dim vcolorsV() As color
    Dim pcolorsV() As color
    
    'Don't duplicate empty groups
    If (obj.Groups(g_index).numvert > 0 And obj.Groups(g_index).numPoly > 0) Then
        With obj.Groups(g_index)
            ReDim vertsV(.numvert - 1)
            ReDim facesV(.numPoly - 1)
            If (.texFlag = 1) Then _
                ReDim TexCoordsV(.numvert - 1)
            ReDim vcolorsV(.numvert - 1)
            ReDim pcolorsV(.numPoly - 1)
            
            CopyMemory vertsV(0), obj.Verts(.offvert), .numvert * 3 * 4
            CopyMemory facesV(0), obj.polys(.offpoly), .numPoly * 24
            If (.texFlag = 1) Then _
                CopyMemory TexCoordsV(0), obj.TexCoords(.offTex), .numvert * 2 * 4
            CopyMemory vcolorsV(0), obj.vcolors(.offvert), .numvert * 4
            CopyMemory pcolorsV(0), obj.PColors(.offpoly), .numPoly * 4
        End With
        
        AddGroup obj, vertsV, facesV, TexCoordsV, vcolorsV, pcolorsV
        
        With obj.Groups(obj.head.NumGroups - 1)
            .texID = obj.Groups(g_index).texID
        End With
        
        DuplicateGroup = True
    Else
        DuplicateGroup = False
    End If
End Function

Public Sub MirrorGroupRelativeToPlane(ByRef obj As PModel, ByVal g_index As Integer, _
                                        ByVal a As Single, ByVal B As Single, ByVal C As Single, _
                                        ByVal d As Single)
    Dim vi As Integer
    Dim PI As Single
    Dim aux As Integer
    Dim p_aux As Point3D
    
    With obj.Groups(g_index)
        For vi = .offvert To .offvert + .numvert - 1
            p_aux = GetPointMirroredRelativeToPlane(obj.Verts(vi), a, B, C, d)
            If (CalculateDistance(p_aux, obj.Verts(vi)) > 0.00001) Then _
                obj.Verts(vi) = p_aux
        Next vi
        
        'Flip faces
        For PI = .offpoly To .offpoly + .numPoly - 1
            aux = obj.polys(PI).Verts(0)
            obj.polys(PI).Verts(0) = obj.polys(PI).Verts(1)
            obj.polys(PI).Verts(1) = aux
        Next PI
    End With
End Sub

Public Sub ApplyPModelTransformation(ByRef obj As PModel, ByRef trans_mat() As Double)
    Dim temp_point As Point3D
    
    Dim num_verts As Long
    Dim vi As Long
    
    num_verts = obj.head.NumVerts
    
    For vi = 0 To num_verts - 1
        MultiplyPoint3DByOGLMatrix trans_mat, obj.Verts(vi), temp_point
        obj.Verts(vi) = temp_point
    Next vi
End Sub

Public Sub RotatePModelModifiers(ByRef obj As PModel, ByVal alpha As Single, ByVal Beta As Single, _
                                    ByVal Gamma As Single)
    Dim diff_alpha As Single
    Dim diff_beta As Single
    Dim diff_gamma As Single
    
    Dim aux_quat As Quaternion
    Dim res_quat As Quaternion
    
    With obj
        If (alpha = 0 Or Beta = 0 Or Gamma = 0) Then
            'This works if there are at most 2 active axes
            BuildQuaternionFromEuler alpha, Beta, Gamma, .RotationQuaternion
        Else
            'Else add up the quaternion difference
            diff_alpha = alpha - .RotateAlpha
            diff_beta = Beta - .RotateBeta
            diff_gamma = Gamma - .RotateGamma
            
            BuildQuaternionFromEuler diff_alpha, diff_beta, diff_gamma, aux_quat
            
            MultiplyQuaternions .RotationQuaternion, aux_quat, res_quat
            
            .RotationQuaternion = res_quat
        End If
        
        .RotateAlpha = alpha
        .RotateBeta = Beta
        .RotateGamma = Gamma
    End With
End Sub

Public Sub SmoothPModel(ByRef obj As PModel)
    Dim gi As Integer
    Dim vi As Long
    Dim PI As Long
    
    Dim group_out As PGroup
    Dim polys_out() As PPolygon
    Dim verts_out() As Point3D
    Dim v_colors_out() As color
    Dim tex_coords_out() As Point2D
    
    Dim num_verts As Long
    Dim num_polys As Long
    Dim num_tex_coords As Long
    
    Dim num_o_polys As Long
    Dim num_o_verts As Long
    
    With obj
        ComputeNormals obj
        For gi = 0 To .head.NumGroups - 1
            SmoothPGroup .Groups(gi), .polys, .Verts, .Normals, .vcolors, .TexCoords, group_out, polys_out, verts_out, v_colors_out, tex_coords_out
            CopyMemory .Groups(gi), group_out, 14 * 4
        Next gi
        
        num_verts = UBound(verts_out) + 1
        num_polys = UBound(polys_out) + 1
        
        .head.NumPolys = num_polys
        .head.NumVerts = num_verts
        
        ReDim Preserve .polys(num_polys - 1)
        ReDim Preserve .Verts(num_verts - 1)
        ReDim Preserve .vcolors(num_verts - 1)
        CopyMemory .polys(0), polys_out(0), num_polys * 24
        CopyMemory .Verts(0), verts_out(0), num_verts * 3 * 4
        CopyMemory .vcolors(0), v_colors_out(0), num_verts * 4
        
        If SafeArrayGetDim(tex_coords_out) > 0 Then
            num_tex_coords = UBound(tex_coords_out) + 1
            ReDim .TexCoords(num_tex_coords - 1)
            CopyMemory .TexCoords(0), tex_coords_out(0), num_tex_coords * 2 * 4
        End If
        
        'For vi = 0 To .head.NumVerts - 1
        '    Debug.Print verts_out(vi).x; ", "; verts_out(vi).y; ", "; verts_out(vi).z
        'Next vi
    End With
    
    ComputeBoundingBox obj
    ComputeNormals obj
    ComputePColors obj
    ComputeEdges obj
End Sub

