Attribute VB_Name = "KimeraAuxiliarModule"
Option Explicit
Type pair_i_b
    i As Integer
    B As Single
End Type

Private Const CFG_FILE_NAME = "Kimera.cfg"

Private Const UNDO_BUFFER_CAPACITY_KEY = "UNDO_BUFFER_CAPACITY"
Private Const CHAR_LGP_PATH_KEY = "CHAR_LGP_PATH"
Private Const BATTLE_LGP_PATH_KEY = "BATTLE_LGP_PATH"
Private Const MAGIC_LGP_PATH_KEY = "MAGIC_LGP_PATH"
Private Const CHAR_LGP_PATH_DEST_KEY = "CHAR_LGP_PATH_DEST"
Private Const BATTLE_LGP_PATH_DEST_KEY = "BATTLE_LGP_PATH_DEST"
Private Const MAGIC_LGP_PATH_DEST_KEY = "MAGIC_LGP_PATH_DEST"
Private Const DEFAULT_FIELD_INTERP_FRAMES_KEY = "DEFAULT_FIELD_INTERP_FRAMES"
Private Const DEFAULT_BATTLE_INTERP_FRAMES_KEY = "DEFAULT_BATTLE_INTERP_FRAMES"

Public UNDO_BUFFER_CAPACITY As Integer
Public CHAR_LGP_PATH As String
Public BATTLE_LGP_PATH As String
Public MAGIC_LGP_PATH As String
Public CHAR_LGP_PATH_DEST As String
Public BATTLE_LGP_PATH_DEST As String
Public MAGIC_LGP_PATH_DEST As String
Public DEFAULT_FIELD_INTERP_FRAMES As Integer
Public DEFAULT_BATTLE_INTERP_FRAMES As Integer


Public Const BARRET_BATTLE_SKELETON = "seaa"
Public Const CID_BATTLE_SKELETON = "rzaa"
Public Const CLOUD_BATTLE_SKELETON = "siaa"
Public Const AERITH_BATTLE_SKELETON = "rvaa"
Public Const RED_BATTLE_SKELETON = "rwaa"
Public Const YUFFIE_BATTLE_SKELETON = "rxaa"
Public Const TIFA_BATTLE_SKELETON = "ruaa"
Public Const CAITSITH_BATTLE_SKELETON = "ryaa"
Public Const VINCENT_BATTLE_SKELETON = "sgaa"
Public Const SEPHIROTH_BATTLE_SKELETON = "saaa"
Public Const FROG_BATTLE_SKELETON = "rsaa"

Type char_lpg_register
    filename As String
    NumNames As Integer
    Names() As String
    NumAnims As Integer
    Animations() As String
End Type

Private Const CHAR_LGP_FILTER_FILE_NAME = "ifalna.fil"
Public NumCharLGPRegisters As Integer
Public CharLGPRegisters() As char_lpg_register

Public EditedPModel As PModel
Public EditedGroup As PGroup
Public OGLContext As Long
Public Sub UpdateTranslationTable(ByRef translation_table_vertex() As pair_i_b, ByRef obj As PModel, ByVal p_index As Integer, ByVal c_index As Integer)
    Dim vi As Integer
    Dim Group As Integer
    Dim diff As Integer
    Dim base_vert As Integer
    
    diff = obj.head.NumVerts - 1 - UBound(translation_table_vertex)
    
    Group = GetPolygonGroup(obj.Groups, p_index)
    base_vert = obj.Groups(Group).offvert + obj.Groups(Group).numvert - 1 - diff
    
    ReDim Preserve translation_table_vertex(obj.head.NumVerts - 1)
    
    For vi = obj.head.NumVerts - 1 To base_vert + 1 Step -1
        With translation_table_vertex(vi)
            .i = translation_table_vertex(vi - diff).i
            .B = translation_table_vertex(vi - diff).B
        End With
    Next vi
    
    
    For vi = base_vert + 1 To base_vert + diff
        With translation_table_vertex(vi)
            .i = c_index
            .B = 1
        End With
    Next vi
    
End Sub
Public Sub ApplyColorTable(ByRef obj As PModel, ByRef color_table() As color, ByRef translation_table_vertex() As pair_i_b)
    Dim gi As Integer
    Dim vi As Integer
    
    Dim C As color
    Dim dv As Double
    
    For vi = 0 To obj.head.NumVerts - 1
        
        C = color_table(translation_table_vertex(vi).i)
        dv = translation_table_vertex(vi).B
        
        With obj.vcolors(vi)
            .r = max(0, Min(255, Fix(C.r / dv)))
            .g = max(0, Min(255, Fix(C.g / dv)))
            .B = max(0, Min(255, Fix(C.B / dv)))
        End With
    Next vi
End Sub
Public Sub fill_color_table(ByRef obj As PModel, ByRef color_table() As color, ByRef n_colors As Long, ByRef translation_table_vertex() As pair_i_b, ByRef translation_table_polys() As pair_i_b, ByVal threshold As Byte)
    Dim v As Single
    Dim dv As Double
    Dim temp_r, temp_g, temp_b As Integer
    Dim C As Integer
    Dim col As color
    Dim it As Integer, i As Integer
    Dim diff As Long
    
    With obj.head
        ReDim color_table(.NumVerts + .NumPolys + 1)
        ReDim translation_table_polys(.NumPolys - 1)
        ReDim translation_table_vertex(.NumVerts - 1)
    End With
    
    For it = 0 To obj.head.NumVerts - 1
        col = obj.vcolors(it)
        
        v = getBrightness(col.r, col.g, col.B)
        ''Debug.Print "Brightness(" + Str$(it) + "):" + Str$(v)
    
        If v = 0 Then
            dv = 255
        Else
            dv = 128 / v
        End If
        temp_r = Min(255, Fix(col.r * dv))
        temp_g = Min(255, Fix(col.g * dv))
        temp_b = Min(255, Fix(col.B * dv))
        C = -1
        diff = 765
        
        For i = 0 To n_colors - 1
            With color_table(i)
                If (.r <= Min(255, temp_r + threshold) And _
                    .r >= max(0, temp_r - threshold)) And _
                   (.g <= Min(255, temp_g + threshold) And _
                    .g >= max(0, temp_g - threshold)) And _
                   (.B <= Min(255, temp_b + threshold) And _
                    .B >= max(0, temp_b - threshold)) Then
                        If Abs(temp_r - .r) + _
                           Abs(temp_g - .g) + _
                           Abs(temp_b - .B) < diff Then
                           
                            diff = Abs(temp_r - .r) + _
                                   Abs(temp_g - .g) + _
                                   Abs(temp_b - .B)
                            C = i
                        End If
                    End If
            End With
        Next i
    
        If C = -1 Then
            color_table(n_colors).r = temp_r
            color_table(n_colors).g = temp_g
            color_table(n_colors).B = temp_b

            n_colors = n_colors + 1
        End If
            
        If C = -1 Then C = n_colors - 1
        
        With translation_table_vertex(it)
            translation_table_vertex(it).i = C
            translation_table_vertex(it).B = dv
        End With
    Next it
    
    For it = 0 To obj.head.NumPolys - 1
        col = obj.PColors(it)
        
        v = getBrightness(col.r, col.g, col.B)
    
        If v = 0 Then
            dv = 128
        Else
            dv = 128 / v
        End If
        temp_r = Min(255, Fix(col.r * dv))
        temp_g = Min(255, Fix(col.g * dv))
        temp_b = Min(255, Fix(col.B * dv))
        C = -1
        diff = 765
        
        For i = 0 To n_colors - 1
            With color_table(i)
                If (.r <= Min(255, temp_r + threshold) And _
                    .r >= max(0, temp_r - threshold)) And _
                   (.g <= Min(255, temp_g + threshold) And _
                    .g >= max(0, temp_g - threshold)) And _
                   (.B <= Min(255, temp_b + threshold) And _
                    .B >= max(0, temp_b - threshold)) Then
                        If Abs(temp_r - .r) + _
                           Abs(temp_g - .g) + _
                           Abs(temp_b - .B) < diff Then
                           
                            diff = Abs(temp_r - .r) + _
                                   Abs(temp_g - .g) + _
                                   Abs(temp_b - .B)
                            C = i
                        End If
                    End If
            End With
        Next i
    
        If C = -1 Then
            With color_table(n_colors)
                .r = temp_r
                .g = temp_g
                .B = temp_b
            End With
            
            n_colors = n_colors + 1
            C = n_colors - 1
        End If
            
        translation_table_polys(it).i = C
        
        If dv = 0 Then
            translation_table_polys(it).B = 0.001
        Else
            translation_table_polys(it).B = dv
        End If
    Next it
End Sub
Sub ComputeFakeEdges(ByRef obj As PModel)
    Dim gi As Integer
    Dim PI As Integer
    Dim ei As Integer
    Dim vi As Integer
    
    ReDim obj.Edges(obj.head.NumPolys * 3)
    
    Dim num_edges As Integer
    Dim found As Boolean
    
    For gi = 0 To obj.head.NumGroups - 1
        obj.Groups(gi).offEdge = num_edges
        For PI = obj.Groups(gi).offpoly To obj.Groups(gi).offpoly + obj.Groups(gi).numPoly - 1
            For vi = 0 To 2
                found = False
                For ei = obj.Groups(gi).offEdge To num_edges - 1
                    With obj.Edges(ei - obj.Groups(gi).offEdge)
                        If (ComparePoints3D(obj.Verts(.Verts(0)), obj.Verts(obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert)) And _
                            ComparePoints3D(obj.Verts(.Verts(1)), obj.Verts(obj.polys(PI).Verts((vi + 1) Mod 3) + obj.Groups(gi).offvert))) Or _
                           (ComparePoints3D(obj.Verts(.Verts(1)), obj.Verts(obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert)) And _
                            ComparePoints3D(obj.Verts(.Verts(0)), obj.Verts(obj.polys(PI).Verts((vi + 1) Mod 3) + obj.Groups(gi).offvert))) Then
                            found = True
                            Exit For
                        End If
                    End With
                Next ei
                
                If Not found Then
                    With obj.Edges(num_edges)
                        .Verts(0) = obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert
                        .Verts(1) = obj.polys(PI).Verts((vi + 1) Mod 3) + obj.Groups(gi).offvert
                    End With
                    obj.polys(PI).Edges(vi) = num_edges - obj.Groups(gi).offEdge
                    num_edges = num_edges + 1
                Else
                    'If Not (((obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert) = obj.Edges(ei - obj.Groups(gi).offEdge).Verts(0) And _
                    '   (obj.polys(PI).Verts((vi + 1) Mod 3) + obj.Groups(gi).offvert) = obj.Edges(ei - obj.Groups(gi).offEdge).Verts(1)) Or _
                    '   ((obj.polys(PI).Verts(vi) + obj.Groups(gi).offvert) = obj.Edges(ei - obj.Groups(gi).offEdge).Verts(1) And _
                    '   (obj.polys(PI).Verts((vi + 1) Mod 3) + obj.Groups(gi).offvert) = obj.Edges(ei - obj.Groups(gi).offEdge).Verts(0))) Then _
                    '    Debug.Print obj.Edges(ei - obj.Groups(gi).offEdge).Verts(0)
                    obj.polys(PI).Edges(vi) = ei - obj.Groups(gi).offEdge
                End If
            Next vi
        Next PI
        
        obj.Groups(gi).numEdge = num_edges - obj.Groups(gi).offEdge
    Next gi
    
    obj.head.NumEdges = num_edges
    'ReDim Preserve obj.Edges(num_edges)
End Sub

Sub DrawFakeEdges(ByRef obj As PModel)
    Dim ei As Integer
    Dim vi As Integer
    
    glColor3f 0, 0, 0
    For ei = 0 To obj.head.NumEdges - 1
        glBegin GL_LINES
            For vi = 0 To 1
                With obj.Verts(obj.Edges(ei).Verts(vi))
                    glVertex3f .x, .y, .z
                End With
            Next vi
        glEnd
    Next ei
End Sub
'Split a polygon through one of it's edges given the alpha parameter for the line equation. Return the new vertex index.
Sub CutEdge(ByRef obj As PModel, ByVal p_index As Integer, ByVal e_index As Integer, ByRef alpha As Single, ByRef intersection_vert As Integer)
    Dim Group As Integer
    
    Dim vi1 As Integer
    Dim vi2 As Integer
    Dim vi3 As Integer
    Dim vi_new As Integer
    Dim tci1 As Integer
    Dim tci2 As Integer
    
    Dim v_buff1(2) As Integer
    Dim v_buff2(2) As Integer
    
    Dim intersection_point As Point3D
    Dim intersection_tex_coord As Point3D
    
    Dim tc1 As Point3D
    Dim tc2 As Point3D
    
    Dim col_temp As color
    
    Group = GetPolygonGroup(obj.Groups, p_index)
    
    With obj.polys(p_index)
        vi1 = .Verts(0) + obj.Groups(Group).offvert
        vi2 = .Verts(1) + obj.Groups(Group).offvert
        vi3 = .Verts(2) + obj.Groups(Group).offvert
    End With
    
    With obj.polys(p_index)
        Select Case e_index
            Case 0:
                col_temp = CombineColor(obj.vcolors(vi1), obj.vcolors(vi2))
            Case 1:
                col_temp = CombineColor(obj.vcolors(vi2), obj.vcolors(vi3))
            Case 2:
                col_temp = CombineColor(obj.vcolors(vi3), obj.vcolors(vi1))
        End Select
    
        Select Case e_index
            Case 0:
                intersection_point = CalculateLinePoint(alpha, obj.Verts(vi1), obj.Verts(vi2))
                If obj.Groups(Group).texFlag = 1 Then
                    tci1 = .Verts(0)
                    tci2 = .Verts(1)
                End If
                
                vi_new = AddVertex(obj, Group, intersection_point, col_temp)
                v_buff1(0) = vi1
                v_buff1(1) = vi_new
                v_buff1(2) = vi3
                
                v_buff2(2) = vi3
                v_buff2(1) = vi2
                v_buff2(0) = vi_new
            Case 1:
                intersection_point = CalculateLinePoint(alpha, obj.Verts(vi2), obj.Verts(vi3))
                If obj.Groups(Group).texFlag = 1 Then
                    tci1 = .Verts(1)
                    tci2 = .Verts(2)
                End If
                
                vi_new = AddVertex(obj, Group, intersection_point, col_temp)
                v_buff1(0) = vi1
                v_buff1(1) = vi2
                v_buff1(2) = vi_new
                
                v_buff2(2) = vi3
                v_buff2(1) = vi_new
                v_buff2(0) = vi1
            Case 2:
                intersection_point = CalculateLinePoint(alpha, obj.Verts(vi3), obj.Verts(vi1))
                If obj.Groups(Group).texFlag = 1 Then
                    tci1 = .Verts(2)
                    tci2 = .Verts(0)
                End If
                
                vi_new = AddVertex(obj, Group, intersection_point, col_temp)
                v_buff1(0) = vi1
                v_buff1(1) = vi2
                v_buff1(2) = vi_new
                
                v_buff2(2) = vi3
                v_buff2(1) = vi2
                v_buff2(0) = vi_new
        End Select
    End With
    
    RemovePolygon obj, p_index
    AddPolygon obj, v_buff1
    AddPolygon obj, v_buff2
    
    If obj.Groups(Group).texFlag = 1 Then
        With obj.TexCoords(tci1 + obj.Groups(Group).offTex)
            tc1.x = .x
            tc1.y = .y
            tc1.z = 0
        End With
        
        With obj.TexCoords(tci2 + obj.Groups(Group).offTex)
            tc2.x = .x
            tc2.y = .y
            tc2.z = 0
        End With
        
        intersection_tex_coord = CalculateLinePoint(alpha, tc1, tc2)
        
        With obj.TexCoords(obj.Groups(Group).offTex + vi_new - obj.Groups(Group).offvert)
            .x = intersection_tex_coord.x
            .y = intersection_tex_coord.y
        End With
    End If
    intersection_vert = vi_new
End Sub
'Split a polygon through one of it's edges given a point and a tex_coord (dummy if the group is untextured)
'Must notify wether the edge was actually cut or not (if the cut point was on one of the vertices)
Function CutEdgeAtPoint(ByRef obj As PModel, ByVal p_index As Integer, ByVal e_index As Integer, ByRef intersection_point As Point3D, ByRef intersection_tex_coord As Point2D) As Boolean
    Dim Group As Integer
    
    Dim vi1 As Integer
    Dim vi2 As Integer
    Dim vi3 As Integer
    Dim vi_new As Integer
    
    Dim v_buff1(2) As Integer
    Dim v_buff2(2) As Integer
    
    Dim col_temp As color
    
    Group = GetPolygonGroup(obj.Groups, p_index)
    
    With obj.polys(p_index)
        vi1 = .Verts(0) + obj.Groups(Group).offvert
        vi2 = .Verts(1) + obj.Groups(Group).offvert
        vi3 = .Verts(2) + obj.Groups(Group).offvert
    End With
    
    With obj.polys(p_index)
        CutEdgeAtPoint = False
        Select Case e_index
            Case 0:
                'It makes no sens cutting an edge through one of it's vertices)
                If ComparePoints3D(obj.Verts(vi1), intersection_point) Or _
                    ComparePoints3D(obj.Verts(vi2), intersection_point) Then _
                        Exit Function
                col_temp = CombineColor(obj.vcolors(vi1), obj.vcolors(vi2))
            Case 1:
                If ComparePoints3D(obj.Verts(vi2), intersection_point) Or _
                    ComparePoints3D(obj.Verts(vi3), intersection_point) Then _
                        Exit Function
                col_temp = CombineColor(obj.vcolors(vi2), obj.vcolors(vi3))
            Case 2:
                If ComparePoints3D(obj.Verts(vi3), intersection_point) Or _
                    ComparePoints3D(obj.Verts(vi1), intersection_point) Then _
                        Exit Function
                col_temp = CombineColor(obj.vcolors(vi3), obj.vcolors(vi1))
        End Select
        
        vi_new = AddVertex(obj, Group, intersection_point, col_temp)
    
        Select Case e_index
            Case 0:
                v_buff1(0) = vi1
                v_buff1(1) = vi_new
                v_buff1(2) = vi3
                
                v_buff2(2) = vi3
                v_buff2(1) = vi2
                v_buff2(0) = vi_new
            Case 1:
                v_buff1(0) = vi1
                v_buff1(1) = vi2
                v_buff1(2) = vi_new
                
                v_buff2(2) = vi3
                v_buff2(1) = vi_new
                v_buff2(0) = vi1
            Case 2:
                v_buff1(0) = vi1
                v_buff1(1) = vi2
                v_buff1(2) = vi_new
                
                v_buff2(2) = vi3
                v_buff2(1) = vi2
                v_buff2(0) = vi_new
        End Select
    End With
    
    RemovePolygon obj, p_index
    AddPolygon obj, v_buff1
    AddPolygon obj, v_buff2
    
    If obj.Groups(Group).texFlag = 1 Then
        With obj.TexCoords(obj.Groups(Group).offTex + vi_new - obj.Groups(Group).offvert)
            .x = intersection_tex_coord.x
            .y = intersection_tex_coord.y
        End With
    End If
    
    CutEdgeAtPoint = True
End Function
Sub CutEdgeAt(ByRef obj As PModel, ByVal p_index As Integer, ByVal e_index As Integer, ByVal intersection_vert As Integer)
    Dim Group As Integer
    
    Dim vi1 As Integer
    Dim vi2 As Integer
    Dim vi3 As Integer
    Dim vi_new As Integer
    Dim tci1 As Integer
    Dim tci2 As Integer
    
    Dim v_buff1(2) As Integer
    Dim v_buff2(2) As Integer
    
    Dim tc1 As Point3D
    Dim tc2 As Point3D
    
    Dim col_temp As color
    
    Group = GetPolygonGroup(obj.Groups, p_index)
    
    With obj.polys(p_index)
        vi1 = .Verts(0) + obj.Groups(Group).offvert
        vi2 = .Verts(1) + obj.Groups(Group).offvert
        vi3 = .Verts(2) + obj.Groups(Group).offvert
    End With
    
    vi_new = intersection_vert
    
    With obj.polys(p_index)
        Select Case e_index
            Case 0:
                col_temp = CombineColor(obj.vcolors(vi1), obj.vcolors(vi2))
            Case 1:
                col_temp = CombineColor(obj.vcolors(vi2), obj.vcolors(vi3))
            Case 2:
                col_temp = CombineColor(obj.vcolors(vi3), obj.vcolors(vi1))
        End Select
    
        Select Case e_index
            Case 0:
                'intersection_point = CalculateLinePoint(Alpha, obj.Verts(vi1), obj.Verts(vi2))
                If obj.Groups(Group).texFlag = 1 Then
                    tci1 = .Verts(0)
                    tci2 = .Verts(1)
                End If
                
                'vi_new = AddVertex(obj, Group, intersection_point, col_temp)
                v_buff1(0) = vi1
                v_buff1(1) = vi_new
                v_buff1(2) = vi3
                
                v_buff2(2) = vi3
                v_buff2(1) = vi2
                v_buff2(0) = vi_new
            Case 1:
                'intersection_point = CalculateLinePoint(Alpha, obj.Verts(vi2), obj.Verts(vi3))
                If obj.Groups(Group).texFlag = 1 Then
                    tci1 = .Verts(1)
                    tci2 = .Verts(2)
                End If
                
                'vi_new = AddVertex(obj, Group, intersection_point, col_temp)
                v_buff1(0) = vi1
                v_buff1(1) = vi2
                v_buff1(2) = vi_new
                
                v_buff2(2) = vi3
                v_buff2(1) = vi_new
                v_buff2(0) = vi1
            Case 2:
                'intersection_point = CalculateLinePoint(Alpha, obj.Verts(vi3), obj.Verts(vi1))
                If obj.Groups(Group).texFlag = 1 Then
                    tci1 = .Verts(2)
                    tci2 = .Verts(0)
                End If
                
                'vi_new = AddVertex(obj, Group, intersection_point, col_temp)
                v_buff1(0) = vi1
                v_buff1(1) = vi2
                v_buff1(2) = vi_new
                
                v_buff2(2) = vi3
                v_buff2(1) = vi2
                v_buff2(0) = vi_new
        End Select
        ''Debug.Print v_buff2(0), v_buff2(1), v_buff2(2), vi_new
    End With
    
    RemovePolygon obj, p_index
    AddPolygon obj, v_buff1
    AddPolygon obj, v_buff2
End Sub

Sub OrderVertices(ByRef obj As PModel, ByRef v_buff() As Integer)
    Dim v1 As Point3D, v2 As Point3D, v3 As Point3D
    Dim aux As Integer
    
    'glMatrixMode GL_MODELVIEW
    'glPushMatrix
    'With obj
    '    glScalef .ResizeX, .ResizeY, .ResizeZ
    '    glRotatef .RotateAlpha, 1, 0, 0
    '    glRotatef .RotateBeta, 0, 1, 0
    '    glRotatef .RotateGamma, 0, 0, 1
    '    glTranslatef .RepositionX, .RepositionY, .RepositionZ
    'End With
    
    v1 = GetVertexProjectedCoords(obj.Verts, v_buff(0))
    v2 = GetVertexProjectedCoords(obj.Verts, v_buff(1))
    v3 = GetVertexProjectedCoords(obj.Verts, v_buff(2))


    If CalculateNormal(v1, v2, v3).z > 0 Then
        aux = v_buff(0)
        v_buff(0) = v_buff(1)
        v_buff(1) = aux
        If CalculateNormal(v2, v1, v3).z > 0 Then
            aux = v_buff(1)
            v_buff(1) = v_buff(2)
            v_buff(2) = aux
            If CalculateNormal(v2, v3, v1).z > 0 Then
                aux = v_buff(0)
                v_buff(0) = v_buff(1)
                v_buff(1) = aux
                If CalculateNormal(v3, v2, v1).z > 0 Then
                    aux = v_buff(1)
                    v_buff(1) = v_buff(2)
                    v_buff(2) = aux
                    If CalculateNormal(v3, v1, v2).z > 0 Then
                        aux = v_buff(0)
                        v_buff(0) = v_buff(1)
                        v_buff(1) = aux
                    End If
                End If
            End If
        End If
    End If
    
    'glMatrixMode GL_MODELVIEW
    'glPopMatrix
End Sub
Public Sub SetLighting(ByVal LightNumber As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal infintyFarQ As Boolean)
    Dim l_color(4) As Single
    Dim l_pos(4) As Single
    
    l_pos(0) = x
    l_pos(1) = y
    l_pos(2) = z
    l_pos(3) = IIf(infintyFarQ, 0, 1)
    
    l_color(0) = red
    l_color(1) = green
    l_color(2) = blue
    l_color(3) = 1
    
    glEnable GL_LIGHTING
    glDisable LightNumber
    
    glLightfv LightNumber, GL_POSITION, l_pos(0)
    glLightfv LightNumber, GL_DIFFUSE, l_color(0)
    glEnable LightNumber
End Sub
Public Sub Fatten(ByRef obj As PModel)
    Dim vi As Integer
    
    Dim CentralZ As Single
    Dim diff_max As Double
    Dim diff_min As Double
    Dim factor As Single
    
    With obj.BoundingBox
        CentralZ = 0
        diff_max = Abs(.max_z - CentralZ)
        diff_min = Abs(CentralZ - .min_z)
    End With
    
    For vi = 0 To obj.head.NumVerts - 1
        With obj.Verts(vi)
            If .z > CentralZ Then
                If diff_max = 0 Then
                    factor = 1
                Else
                    factor = 1 + (1 - Abs(.z - CentralZ) / diff_max) * 0.1
                End If
            Else
                If diff_min = 0 Then
                    factor = 1
                Else
                    factor = 1 + (1 - Abs(CentralZ - .z) / diff_min) * 0.1
                End If
            End If
            .x = .x * factor
            .y = .y * factor
        End With
    Next vi
End Sub
Public Sub Slim(ByRef obj As PModel)
    Dim vi As Integer
    
    Dim CentralZ As Single
    Dim diff_max As Double
    Dim diff_min As Double
    Dim factor As Single
    
    With obj.BoundingBox
        CentralZ = 0
        diff_max = Abs(.max_z - CentralZ)
        diff_min = Abs(CentralZ - .min_z)
    End With
    
    For vi = 0 To obj.head.NumVerts - 1
        With obj.Verts(vi)
            If .z > CentralZ Then
                If diff_max = 0 Then
                    factor = 1
                Else
                    factor = 1 + (1 - Abs(.z - CentralZ) / diff_max) * 0.1
                End If
            Else
                If diff_min = 0 Then
                    factor = 1
                Else
                    factor = 1 + (1 - Abs(CentralZ - .z) / diff_min) * 0.1
                End If
            End If
            .x = .x / factor
            .y = .y / factor
        End With
    Next vi
End Sub
Public Sub RemoveTexturedGroups(ByRef obj As PModel)
    Dim gi As Integer
    
    With obj
        For gi = .head.NumGroups - 1 To 0 Step -1
            If .Groups(gi).texFlag = 1 Then _
                RemoveGroup obj, gi
        Next gi
    End With
End Sub
Public Sub HorizontalMirror(ByRef obj As PModel)
    Dim vi As Integer
    Dim PI As Integer
    
    For vi = 0 To obj.head.NumVerts - 1
        With obj.Verts(vi)
            .x = -1 * .x
        End With
    Next vi
    
    For PI = 0 To obj.head.NumPolys - 1
        With obj.polys(PI)
            vi = .Verts(1)
            .Verts(1) = .Verts(0)
            .Verts(0) = vi
        End With
    Next PI
End Sub
Public Sub ChangeBrigthness(ByRef obj As PModel, ByVal factor As Integer)
    Dim vi As Integer
    Dim PI As Integer
    
    For vi = 0 To obj.head.NumVerts - 1
        With obj.vcolors(vi)
            .r = max(0, Min(255, .r + factor))
            .g = max(0, Min(255, .g + factor))
            .B = max(0, Min(255, .B + factor))
        End With
    Next vi
    
    For PI = 0 To obj.head.NumPolys - 1
        With obj.PColors(PI)
            .r = max(0, Min(255, .r + factor))
            .g = max(0, Min(255, .g + factor))
            .B = max(0, Min(255, .B + factor))
        End With
    Next PI
End Sub
Public Sub KillPrecalculatedLighting(ByRef obj As PModel, ByRef translation_table_vertex() As pair_i_b)
    Dim ci As Integer
    
    For ci = 0 To obj.head.NumVerts - 1
        translation_table_vertex(ci).B = 1
    Next ci
End Sub
Public Function IsCameraUnderGround() As Boolean
    Dim origin As Point3D
    Dim origin_trans As Point3D
    Dim MV_matrix(16) As Double
    
    glGetDoublev GL_MODELVIEW_MATRIX, MV_matrix(0)
    
    InvertMatrix MV_matrix
    
    With origin
        .x = 0
        .y = 0
        .z = 0
    End With
    
    MultiplyPoint3DByOGLMatrix MV_matrix, origin, origin_trans
    
    IsCameraUnderGround = origin_trans.y > -1
End Function
Public Function GetCameraUnderGroundValue() As Single
    Dim origin As Point3D
    Dim origin_trans As Point3D
    Dim MV_matrix(16) As Double
    
    glGetDoublev GL_MODELVIEW_MATRIX, MV_matrix(0)
    
    InvertMatrix MV_matrix
    
    With origin
        .x = 0
        .y = 0
        .z = 0
    End With
    
    MultiplyPoint3DByOGLMatrix MV_matrix, origin, origin_trans
    
    GetCameraUnderGroundValue = -1 - origin_trans.y
End Function

Public Sub ReadCFGFile()
    Dim line As String
    Dim NFileAux As Integer
    
    On Error GoTo ErrHand
    NFileAux = FreeFile
    Open App.Path + "\" + CFG_FILE_NAME For Input As #NFileAux
    Line Input #NFileAux, line
    UNDO_BUFFER_CAPACITY = val(Right$(line, Len(line) - Len(UNDO_BUFFER_CAPACITY_KEY) - 3))
    Line Input #NFileAux, line
    CHAR_LGP_PATH = Right$(line, Len(line) - Len(CHAR_LGP_PATH_KEY) - 3)
    Line Input #NFileAux, line
    BATTLE_LGP_PATH = Right$(line, Len(line) - Len(BATTLE_LGP_PATH_KEY) - 3)
    Line Input #NFileAux, line
    MAGIC_LGP_PATH = Right$(line, Len(line) - Len(MAGIC_LGP_PATH_KEY) - 3)
    Line Input #NFileAux, line
    CHAR_LGP_PATH_DEST = Right$(line, Len(line) - Len(CHAR_LGP_PATH_DEST_KEY) - 3)
    Line Input #NFileAux, line
    BATTLE_LGP_PATH_DEST = Right$(line, Len(line) - Len(BATTLE_LGP_PATH_DEST_KEY) - 3)
    Line Input #NFileAux, line
    MAGIC_LGP_PATH_DEST = Right$(line, Len(line) - Len(MAGIC_LGP_PATH_DEST_KEY) - 3)
    Line Input #NFileAux, line
    DEFAULT_FIELD_INTERP_FRAMES = val(Right$(line, Len(line) - Len(DEFAULT_FIELD_INTERP_FRAMES_KEY) - 3))
    Line Input #NFileAux, line
    DEFAULT_BATTLE_INTERP_FRAMES = val(Right$(line, Len(line) - Len(DEFAULT_BATTLE_INTERP_FRAMES_KEY) - 3))
    Close #NFileAux
    Exit Sub
ErrHand:
    MsgBox "Error " + Str$(Err) + " reading " + CFG_FILE_NAME + "!!!", vbOKOnly, "Error reading"
    Close #NFileAux
End Sub

Public Sub WriteCFGFile()
    Dim line As String
    Dim NFileAux As Integer
    
    On Error GoTo ErrHand
    NFileAux = FreeFile
    Open App.Path + "\" + CFG_FILE_NAME For Output As #NFileAux
    line = UNDO_BUFFER_CAPACITY_KEY + " = " + Str$(UNDO_BUFFER_CAPACITY)
    Print #NFileAux, line
    line = CHAR_LGP_PATH_KEY + " = " + CHAR_LGP_PATH
    Print #NFileAux, line
    line = BATTLE_LGP_PATH_KEY + " = " + BATTLE_LGP_PATH
    Print #NFileAux, line
    line = MAGIC_LGP_PATH_KEY + " = " + MAGIC_LGP_PATH
    Print #NFileAux, line
    line = CHAR_LGP_PATH_DEST_KEY + " = " + CHAR_LGP_PATH_DEST
    Print #NFileAux, line
    line = BATTLE_LGP_PATH_DEST_KEY + " = " + BATTLE_LGP_PATH_DEST
    Print #NFileAux, line
    line = MAGIC_LGP_PATH_DEST_KEY + " = " + MAGIC_LGP_PATH_DEST
    Print #NFileAux, line
    line = DEFAULT_FIELD_INTERP_FRAMES_KEY + " = " + Str$(DEFAULT_FIELD_INTERP_FRAMES)
    Print #NFileAux, line
    line = DEFAULT_BATTLE_INTERP_FRAMES_KEY + " = " + Str$(DEFAULT_BATTLE_INTERP_FRAMES)
    Print #NFileAux, line
    Close #NFileAux
    Exit Sub
ErrHand:
    MsgBox "Error " + Str$(Err) + " writting " + CFG_FILE_NAME + "!!!", vbOKOnly, "Error writting"
    Close #NFileAux
End Sub
Public Sub ReadCharFilterFile()
    Dim line As String
    Dim NFileAux As Integer
    Dim file_name As String
    Dim last_file_name As String
    Dim name As String
    Dim key As String
    Dim p_data_start As Long
    Dim p_data_end As Long
    Dim line_length As Long
    
    On Error GoTo ErrHand
    NFileAux = FreeFile
    name = ""
    Open App.Path + "\" + CHAR_LGP_FILTER_FILE_NAME For Input As #NFileAux
    Do
        Line Input #NFileAux, line
        line_length = Len(line)
        
        If line_length > 0 Then
            file_name = Left$(line, 4)
            
            If Not last_file_name = file_name Then
                NumCharLGPRegisters = NumCharLGPRegisters + 1
                ReDim Preserve CharLGPRegisters(NumCharLGPRegisters - 1)
                CharLGPRegisters(NumCharLGPRegisters - 1).filename = file_name
                CharLGPRegisters(NumCharLGPRegisters - 1).NumAnims = 0
                CharLGPRegisters(NumCharLGPRegisters - 1).NumNames = 0
            End If
            
            key = Mid$(line, 5, 5)
            
            p_data_start = InStr(8, line, "=") + 1
            With CharLGPRegisters(NumCharLGPRegisters - 1)
                If key = "Names" Then
                    Do
                        p_data_end = InStr(p_data_start, line, ",")
                        ReDim Preserve .Names(.NumNames)
                        .Names(.NumNames) = Mid$(line, p_data_start, p_data_end - p_data_start)
                        .NumNames = .NumNames + 1
                        p_data_start = p_data_end + 1
                    Loop Until p_data_end = line_length
                ElseIf key = "Anims" Then
                    Do
                        p_data_end = InStr(p_data_start, line, ",")
                        ReDim Preserve .Animations(.NumAnims)
                        .Animations(.NumAnims) = Mid$(line, p_data_start, p_data_end - p_data_start)
                        .NumAnims = .NumAnims + 1
                        p_data_start = p_data_end + 1
                    Loop Until p_data_end = line_length
                End If
            End With
            
            last_file_name = file_name
        End If
    Loop Until line = "" Or EOF(NFileAux)
    Close #NFileAux
    Exit Sub
ErrHand:
    MsgBox "Error " + Str$(Err) + " reading " + CHAR_LGP_FILTER_FILE_NAME + "!!!", vbOKOnly, "Error reading"
    Close #NFileAux
End Sub
