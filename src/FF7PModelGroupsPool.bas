Attribute VB_Name = "FF7PModelGroupsPool"
Option Explicit
Type PGroup
    polyType As Long
    offpoly As Long
    numPoly As Long
    offvert As Long
    numvert As Long
    offEdge As Long
    numEdge As Long
    off1c As Long
    off20 As Long
    off24 As Long
    off28 As Long
    offTex As Long
    texFlag As Long
    TexID As Long
'-------------Extra Atributes----------------
    DListNum As Long
    HiddenQ As Boolean  'Hidden groups aren't rendered and can't be changed _
                        save for the basic geometrical transformations (rotation, scaling and panning),
                        'palletizzed opeartions and group deletion
End Type

Private Type VertsGroup
    indices() As Long
    OriginalPosition As Point3D
    normal As Point3D
End Type

Private Const MIN_SMOOTH_COS As Double = -0.2
Sub ReadGroups(ByVal NFile As Integer, ByVal offset As Long, ByRef Groups() As PGroup, ByVal NumGroups As Long)
    Dim gi As Long
    
    ReDim Groups(NumGroups - 1)
    For gi = 0 To NumGroups - 1
        With Groups(gi)
            Get NFile, offset + gi * 56, .polyType
            Get NFile, offset + gi * 56 + 4, .offpoly
            Get NFile, offset + gi * 56 + 8, .numPoly
            Get NFile, offset + gi * 56 + &HC, .offvert
            Get NFile, offset + gi * 56 + &H10, .numvert
            Get NFile, offset + gi * 56 + &H14, .offEdge
            Get NFile, offset + gi * 56 + &H18, .numEdge
            Get NFile, offset + gi * 56 + &H1C, .off1c
            Get NFile, offset + gi * 56 + &H20, .off20
            Get NFile, offset + gi * 56 + &H24, .off24
            Get NFile, offset + gi * 56 + &H28, .off28
            Get NFile, offset + gi * 56 + &H2C, .offTex
            Get NFile, offset + gi * 56 + &H30, .texFlag
            Get NFile, offset + gi * 56 + &H34, .TexID
            .DListNum = -1
            .HiddenQ = False
        End With
    Next gi
End Sub
Sub WriteGroups(ByVal NFile As Integer, ByVal offset As Long, ByRef Groups() As PGroup)
    Dim gi As Long
    Dim NumGroups As Long
    
    NumGroups = UBound(Groups()) + 1
    For gi = 0 To NumGroups - 1
        With Groups(gi)
            Put NFile, offset + gi * 56, .polyType
            Put NFile, offset + gi * 56 + 4, .offpoly
            Put NFile, offset + gi * 56 + 8, .numPoly
            Put NFile, offset + gi * 56 + &HC, .offvert
            Put NFile, offset + gi * 56 + &H10, .numvert
            Put NFile, offset + gi * 56 + &H14, .offEdge
            Put NFile, offset + gi * 56 + &H18, .numEdge
            Put NFile, offset + gi * 56 + &H1C, .off1c
            Put NFile, offset + gi * 56 + &H20, .off20
            Put NFile, offset + gi * 56 + &H24, .off24
            Put NFile, offset + gi * 56 + &H28, .off28
            Put NFile, offset + gi * 56 + &H2C, .offTex
            Put NFile, offset + gi * 56 + &H30, .texFlag
            Put NFile, offset + gi * 56 + &H34, .TexID
        End With
    Next gi
End Sub
Sub MergeGroups(ByRef g1() As PGroup, ByRef g2() As PGroup)
    Dim gi As Integer
    Dim MaxTIG1 As Integer
    Dim NumGroupsG1 As Integer
    Dim NumGroupsG2 As Integer
    Dim NumPolys As Integer
    Dim NumEdges As Integer
    Dim NumVerts As Integer
    Dim NumTexCs As Integer
    
    
    NumGroupsG1 = UBound(g1) + 1
    NumGroupsG2 = UBound(g2) + 1
    
    ReDim Preserve g1(NumGroupsG1 + NumGroupsG2 - 1)
    
    MaxTIG1 = 0
    For gi = 0 To NumGroupsG1 - 1
        If g1(gi).texFlag = 1 Then _
            If g1(gi).TexID > MaxTIG1 Then MaxTIG1 = g1(gi).TexID
    Next gi
    
    With g1(NumGroupsG1 - 1)
        NumPolys = .offpoly + .numPoly
        NumEdges = .offEdge + .numEdge
        NumVerts = .offvert + .numvert
    
        If g1(NumGroupsG1).texFlag = 1 Then
            NumTexCs = .offTex + .numvert
        Else
            NumTexCs = .offTex
        End If
    End With
    
    For gi = 0 To NumGroupsG2 - 1
        With g2(gi)
            .offpoly = .offpoly + NumPolys
            .offvert = .offvert + NumVerts
            .offEdge = .offEdge + NumEdges
            .offTex = .offTex + NumTexCs
            If .texFlag = 1 Then .TexID = .TexID + MaxTIG1
        End With
        g1(NumGroupsG1 + gi) = g2(gi)
    Next gi
End Sub
Sub CreateDListFromPGroup(ByRef Group As PGroup, ByRef polys() As PPolygon, ByRef Verts() As Point3D, ByRef vcolors() As color, ByRef Normals() As Point3D, ByRef TexCoords() As Point2D, ByRef hundret As PHundret)
    With Group
        If .DListNum < 0 Then
            .DListNum = glGenLists(1)
        Else
            glDeleteLists .DListNum, 1
            .DListNum = glGenLists(1)
        End If
        
        glNewList .DListNum, GL_COMPILE
            DrawGroup Group, polys, Verts, vcolors, Normals, TexCoords, hundret, False
        glEndList
    End With
End Sub
Sub FreeGroupResources(ByRef obj As PGroup)
    glDeleteLists obj.DListNum, 1
End Sub
Sub DrawGroupDList(ByRef Group As PGroup)
    glCallList Group.DListNum
End Sub
Sub DrawGroup(ByRef Group As PGroup, ByRef polys() As PPolygon, ByRef Verts() As Point3D, ByRef vcolors() As color, ByRef Normals() As Point3D, _
              ByRef TexCoords() As Point2D, ByRef hundret As PHundret, ByVal HideHiddenQ As Boolean)
    If Group.HiddenQ And HideHiddenQ Then _
        Exit Sub
        
    Dim PI As Integer
    Dim vi As Integer
    Dim TexEnabled As Boolean
    Dim x As Single, y As Single, z As Single
    
    TexEnabled = (glIsEnabled(GL_TEXTURE_2D) = GL_TRUE)
    
    glBegin GL_TRIANGLES
    glColorMaterial GL_FRONT_AND_BACK, GL_AMBIENT_AND_DIFFUSE
    For PI = Group.offpoly To Group.offpoly + Group.numPoly - 1
        
            For vi = 0 To 2
                With vcolors(polys(PI).Verts(vi) + Group.offvert)
                    If hundret.blend_mode = 0 And Not TexEnabled Then
                        glColor4f CSng(.r) / 255#, CSng(.g) / 255#, CSng(.B) / 255#, 0.5  'CSng(.a) / 255# '0.5
                    Else
                        glColor4f CSng(.r) / 255#, CSng(.g) / 255#, CSng(.B) / 255#, 1#
                    End If
                    
                End With
                
                With Normals(polys(PI).Normals(vi))
                    glNormal3f .x, .y, .z
                End With
                
                If Group.texFlag = 1 Then
                    With TexCoords(Group.offTex + polys(PI).Verts(vi))
                        x = .x
                        y = .y
                        glTexCoord2f x, y
                        
                    End With
                End If

                With Verts(polys(PI).Verts(vi) + Group.offvert)
                    x = .x
                    y = .y
                    z = .z
                    glVertex3f x, y, z
                End With
                
            Next vi
        
    Next PI
    glEnd
    
End Sub
Function GetPolygonGroup(ByRef Groups() As PGroup, ByVal PI As Integer) As Integer
    Dim NumGroups As Long
    Dim p_base As Integer
    
    NumGroups = UBound(Groups()) - 1
    
    GetPolygonGroup = 0
    p_base = p_base + Groups(0).numPoly
    
    While p_base <= PI
        GetPolygonGroup = GetPolygonGroup + 1
        p_base = p_base + Groups(GetPolygonGroup).numPoly
    Wend
End Function
Function GetVertexGroup(ByRef Groups() As PGroup, ByVal vi As Integer) As Integer
    Dim NumGroups As Long
    Dim v_base As Integer
    
    NumGroups = UBound(Groups()) - 1
    
    GetVertexGroup = 0
    v_base = v_base + Groups(0).numvert
    
    While v_base <= vi
        GetVertexGroup = GetVertexGroup + 1
        v_base = v_base + Groups(GetVertexGroup).numvert
    Wend
End Function
'Appends to the output the group with a single iteration of Doo-Sabin smoothing applied
Public Sub SmoothPGroup(ByRef group_in As PGroup, _
                        ByRef polys_in() As PPolygon, ByRef verts_in() As Point3D, ByRef normals_in() As Point3D, ByRef v_colors_in() As color, _
                        ByRef tex_coords_in() As Point2D, _
                        ByRef group_out As PGroup, _
                        ByRef polys_out() As PPolygon, ByRef verts_out() As Point3D, ByRef v_colors_out() As color, ByRef tex_coords_out() As Point2D)
    
    Dim polys_aux() As PPolygon
    Dim verts_aux() As Point3D
    Dim normals_aux() As Point3D
    Dim v_colors_aux() As color
    Dim tex_coords_aux() As Point2D
    Dim v_colors_aux_copy() As color
    Dim tex_coords_aux_copy() As Point2D
    
    Dim per_edge_verts() As VertsGroup
    Dim per_vertex_verts() As VertsGroup
    Dim vert_groups_per_edge() As Long
    
    Dim polys_per_edge() As VertsGroup
    Dim vertex_verts_central_verts() As Long
    Dim per_vertex_verts_polys() As VertsGroup
    Dim per_vertex_coefs() As Point2D
    
    Dim num_verts As Long
    
    GetIsolatedGroup group_in, polys_in, verts_in, v_colors_in, tex_coords_in, _
                     polys_aux, verts_aux, v_colors_aux, tex_coords_aux
    
    SplitSharedVertices polys_aux, verts_aux, normals_aux, v_colors_aux, tex_coords_aux, normals_in
    
    CreateEdgeVerticesTable polys_aux, verts_aux, per_edge_verts
    CreateVertexVerticesTable verts_aux, normals_aux, per_edge_verts, per_vertex_verts
    CreateVertGroupsPerEdgeTable verts_aux, per_edge_verts, per_vertex_verts, vert_groups_per_edge
    
    num_verts = UBound(verts_aux) + 1
    ReDim v_colors_aux_copy(num_verts - 1)
    CopyMemory v_colors_aux_copy(0), v_colors_aux(0), 4 * (num_verts - 1)
    If SafeArrayGetDim(tex_coords_aux) > 0 Then
        ReDim tex_coords_aux_copy(num_verts - 1)
        CopyMemory tex_coords_aux_copy(0), tex_coords_aux(0), 2 * 3 * (num_verts - 1)
    End If
    DooSabinPolysContraction polys_aux, verts_aux, v_colors_aux, tex_coords_aux, per_edge_verts, per_vertex_verts, per_vertex_coefs
    
    ConnectEdgeVertices polys_aux, verts_aux, v_colors_aux, tex_coords_aux, per_edge_verts, polys_per_edge
    ConnectVertexVertices polys_aux, verts_aux, v_colors_aux, tex_coords_aux, v_colors_aux_copy, tex_coords_aux_copy, per_vertex_verts, _
                          vertex_verts_central_verts, per_vertex_verts_polys
    
    FixVertexAtributes polys_aux, verts_aux, normals_aux, v_colors_aux, tex_coords_aux, v_colors_aux_copy, tex_coords_aux, _
                       per_vertex_coefs, per_edge_verts, polys_per_edge, vertex_verts_central_verts, per_vertex_verts_polys, vert_groups_per_edge
    
    AppendIsolatedGroup group_in, polys_aux, verts_aux, v_colors_aux, tex_coords_aux, _
                        group_out, polys_out, verts_out, v_colors_out, tex_coords_out

End Sub
'Returns an isoalted copy of the group components
Private Sub GetIsolatedGroup(ByRef Group As PGroup, _
                             ByRef polys_in() As PPolygon, ByRef verts_in() As Point3D, ByRef v_colors_in() As color, ByRef tex_coords_in() As Point2D, _
                             ByRef polys_out() As PPolygon, ByRef verts_out() As Point3D, ByRef v_colors_out() As color, ByRef tex_coords_out() As Point2D)
    'Dim PI As Long
    'Dim vi As Integer
    
    With Group
        ReDim polys_out(.numPoly - 1)
        ReDim verts_out(.numvert - 1)
        ReDim v_colors_out(.numvert - 1)
        If .texFlag = 1 Then
            ReDim tex_coords_out(.numvert - 1)
        End If
        
        CopyMemory polys_out(0), polys_in(.offpoly), .numPoly * 24
        CopyMemory verts_out(0), verts_in(.offvert), .numvert * 3 * 4
        CopyMemory v_colors_out(0), v_colors_in(.offvert), .numvert * 4
        
        If .texFlag = 1 Then
            CopyMemory tex_coords_out(0), tex_coords_in(.offTex), .numvert * 2 * 4
        End If
        
        'For PI = 0 To .numPoly - 1
        '    For vi = 0 To 2
        '        polys_out(PI).Verts(vi) = polys_out(PI).Verts(vi) - .offvert
        '    Next vi
        'Next PI
    End With
End Sub
'------------------------------------------The following functions work with isolated geometry (hence, no need for the group information--------------------------------------
'Appends a vertex with the specified data
Private Function AppendVertex(ByRef vert As Point3D, ByRef normal As Point3D, ByRef v_color As color, ByRef Verts() As Point3D, ByRef Normals() As Point3D, _
                              ByRef v_colors() As color) As Long
    Dim num_verts As Long
    
    num_verts = UBound(Verts) + 1
    ReDim Preserve Verts(num_verts)
    CopyMemory Verts(num_verts), vert, 3 * 4
    ReDim Preserve Normals(num_verts)
    CopyMemory Normals(num_verts), normal, 3 * 4
    ReDim Preserve v_colors(num_verts)
    CopyMemory v_colors(num_verts), v_color, 4
    
    AppendVertex = num_verts
End Function
'Appends a vertex with the specified data (with texture coords)
Private Function AppendVertexWithTexCoords(ByRef vert As Point3D, ByRef normal As Point3D, ByRef v_color As color, ByRef tex_coord As Point2D, _
                                           ByRef Verts() As Point3D, ByRef Normals() As Point3D, ByRef v_colors() As color, ByRef tex_coords() As Point2D) As Long
    Dim num_verts As Long
    Dim num_tex_coords As Long
    
    num_verts = AppendVertex(vert, normal, v_color, Verts, Normals, v_colors)
    num_tex_coords = UBound(tex_coords) + 1
    ReDim Preserve tex_coords(num_tex_coords)
    CopyMemory tex_coords(num_tex_coords), tex_coord, 2 * 4
    
    AppendVertexWithTexCoords = num_verts
End Function
'Splits all shared vertices
Private Sub SplitSharedVertices(ByRef polys() As PPolygon, ByRef Verts() As Point3D, ByRef Normals() As Point3D, ByRef v_colors() As color, _
                                ByRef tex_coords() As Point2D, ByRef normals_in() As Point3D)
    Dim num_polys As Long
    Dim num_verts As Long
    Dim used_vetrs() As Boolean
    Dim PI As Long
    Dim vi As Integer
    Dim vertex_index As Long
    Dim normal_index As Long
    Dim has_tex_coordsQ As Boolean
    Dim aux_vert As Point3D
    Dim aux_normal As Point3D
    Dim aux_v_color As color
    Dim aux_tex_coord As Point2D
    
    num_polys = UBound(polys) + 1
    num_verts = UBound(Verts) + 1
    ReDim Normals(num_verts - 1)
    ReDim used_verts(num_verts - 1) As Boolean
    ZeroMemory used_verts(0), (num_verts - 1) * 2
    
    has_tex_coordsQ = SafeArrayGetDim(tex_coords) > 0
    
    For PI = 0 To num_polys - 1
        With polys(PI)
            For vi = 0 To 2
                vertex_index = .Verts(vi)
                normal_index = .Normals(vi)
                If used_verts(vertex_index) Then
                    aux_vert = Verts(vertex_index)
                    aux_normal = normals_in(normal_index)
                    aux_v_color = v_colors(vertex_index)
                    If has_tex_coordsQ Then
                        aux_tex_coord = tex_coords(vertex_index)
                        vertex_index = AppendVertexWithTexCoords(aux_vert, aux_normal, aux_v_color, aux_tex_coord, Verts, Normals, v_colors, tex_coords)
                    Else
                        vertex_index = AppendVertex(aux_vert, aux_normal, aux_v_color, Verts, Normals, v_colors)
                    End If
                    
                    .Verts(vi) = vertex_index
                    .Normals(vi) = vertex_index
                Else
                    Normals(vertex_index) = normals_in(normal_index)
                    .Normals(vi) = vertex_index
                    used_verts(vertex_index) = True
                End If
            Next vi
        End With
    Next PI
End Sub
'Creates a table of vertex indices implied on a smoothed edge (in order)
Private Sub CreateEdgeVerticesTable(ByRef polys() As PPolygon, ByRef Verts() As Point3D, ByRef edge_verts_out() As VertsGroup)
    Dim PI As Long
    Dim pi2 As Long
    Dim vi As Long
    Dim vi2 As Long
    
    Dim v_indexA1 As Integer
    Dim v_indexA2 As Integer
    Dim v_indexB1 As Integer
    Dim v_indexB2 As Integer
    
    Dim aux_vertA1 As Point3D
    Dim aux_vertA2 As Point3D
    Dim aux_vertB1 As Point3D
    Dim aux_vertB2 As Point3D
    
    Dim num_polys As Long
    Dim num_edges As Long
    Dim match_foundQ As Boolean
    
    num_polys = UBound(polys) + 1
    num_edges = 0
    
    'We will take adventage of edge indices to store topolgy information. To be discarded later.
    For PI = 0 To num_polys - 1
        With polys(PI)
            .Edges(0) = -1
            .Edges(1) = -1
            .Edges(2) = -1
        End With
    Next PI
    
    For PI = 0 To num_polys - 1
        With polys(PI)
            For vi = 0 To 2
                If .Edges(vi) = -1 Then
                    ReDim Preserve edge_verts_out(num_edges)
                    ReDim edge_verts_out(num_edges).indices(1)
                    
                    v_indexA1 = .Verts(vi)
                    v_indexA2 = .Verts((vi + 1) Mod 3)
                    
                    .Edges(vi) = num_edges
                    edge_verts_out(num_edges).indices(0) = v_indexA1
                    edge_verts_out(num_edges).indices(1) = v_indexA2
                    
                    aux_vertA1 = Verts(v_indexA1)
                    aux_vertA2 = Verts(v_indexA2)
                    
                    match_foundQ = False
                    For pi2 = 0 To num_polys - 1
                        If pi2 <> PI Then
                            For vi2 = 0 To 2
                                If polys(pi2).Edges(vi2) = -1 Then
                                    v_indexB1 = polys(pi2).Verts(vi2)
                                    v_indexB2 = polys(pi2).Verts((vi2 + 1) Mod 3)
                                    
                                    aux_vertB1 = Verts(v_indexB1)
                                    aux_vertB2 = Verts(v_indexB2)
                                    
                                    'Adjacent polygons cross the vertices in oposing directions
                                    If CompareSimilarPoints3D(aux_vertA1, aux_vertB2) And CompareSimilarPoints3D(aux_vertA2, aux_vertB1) Then
                                        ReDim Preserve edge_verts_out(num_edges).indices(3)
                                        edge_verts_out(num_edges).indices(2) = v_indexB1
                                        edge_verts_out(num_edges).indices(3) = v_indexB2
                                        
                                        polys(pi2).Edges(vi2) = num_edges
                                        match_foundQ = True
                                        
                                        Exit For
                                    End If
                                End If
                            Next vi2
                        End If
                        
                        If match_foundQ Then Exit For
                    Next pi2
                    num_edges = num_edges + 1
                End If
            Next vi
        End With
    Next PI
End Sub
'Creates a table of vertex indices implied on a smoothed vertex (in order). CreateEdgeVerticesTable must be called previously to build edges list and fill topology information.
Private Sub CreateVertexVerticesTable(ByRef Verts() As Point3D, ByRef Normals() As Point3D, ByRef edge_verts() As VertsGroup, _
                                      ByRef vertex_verts_out() As VertsGroup)
    Dim vi As Long
    Dim vi2 As Long
    Dim v_index_aux As Long
    Dim v_index_next_aux As Long
    Dim ordered_list_index As Long
    Dim vgi As Long
    Dim ei As Long
    Dim last_edge_index As Long
    Dim next_vector_list_index As Long
    
    Dim num_verts As Long
    Dim num_vert_groups As Long
    Dim num_verts_per_vert As Long
    Dim num_verts_per_edge As Long
    
    Dim num_edges  As Long
    Dim edge_usages() As Integer
    
    Dim match_foundQ As Boolean
    Dim must_reverse_orderQ As Boolean
    Dim aux_must_reverse_orderQ As Boolean
    Dim check_opposite_directionQ As Boolean
    
    'First find equivalent vertices
    num_verts = UBound(Verts) + 1
    num_vert_groups = 0
    For vi = 0 To num_verts - 1
        match_foundQ = False
        For vgi = 0 To num_vert_groups - 1
            With vertex_verts_out(vgi)
                If CompareSimilarPoints3D(Verts(.indices(0)), Verts(vi)) Then
                    num_verts_per_vert = UBound(.indices) + 1
                    ReDim Preserve .indices(num_verts_per_vert)
                    .indices(num_verts_per_vert) = vi
                    .normal.x = .normal.x + Normals(vi).x
                    .normal.y = .normal.y + Normals(vi).y
                    .normal.z = .normal.z + Normals(vi).z
                    
                    match_foundQ = True
                    
                    Exit For
                End If
            End With
        Next vgi
        
        If Not match_foundQ Then
            ReDim Preserve vertex_verts_out(num_vert_groups)
            ReDim Preserve vertex_verts_out(num_vert_groups).indices(0)
            With vertex_verts_out(num_vert_groups)
                .indices(0) = vi
                .OriginalPosition = Verts(vi)
                .normal = Normals(vi)
            End With
            num_vert_groups = num_vert_groups + 1
        End If
    Next vi
    
    For vgi = 0 To num_vert_groups - 1
        With vertex_verts_out(vgi)
            num_verts_per_vert = UBound(.indices) + 1
            .normal.x = .normal.x / num_verts_per_vert
            .normal.y = .normal.y / num_verts_per_vert
            .normal.z = .normal.z / num_verts_per_vert
            .normal = Normalize(.normal)
        End With
    Next vgi
    
    'Order equivalent vertices (acording to edges topology)
    num_edges = UBound(edge_verts) + 1
    ReDim edge_usages(num_edges - 1)
    ZeroMemory edge_usages(0), 2 * num_edges
    For vgi = 0 To num_vert_groups - 1
        v_index_aux = vertex_verts_out(vgi).indices(0)
        ordered_list_index = 1
        num_verts_per_vert = UBound(vertex_verts_out(vgi).indices) + 1
        last_edge_index = -1
        must_reverse_orderQ = False
        check_opposite_directionQ = False
        While ordered_list_index < num_verts_per_vert
            'find vertex in edges
            match_foundQ = False
            For ei = 0 To num_edges - 1
                If edge_usages(ei) < 2 Then
                    With edge_verts(ei)
                        If ei <> last_edge_index Then
                            num_verts_per_edge = UBound(.indices) + 1
                            If num_verts_per_edge = 4 Then
                                next_vector_list_index = -1
                                aux_must_reverse_orderQ = False
                                
                                If .indices(0) = v_index_aux Then
                                    v_index_next_aux = .indices(3)
                                    aux_must_reverse_orderQ = True
                                    match_foundQ = True
                                ElseIf .indices(1) = v_index_aux Then
                                    v_index_next_aux = .indices(2)
                                    match_foundQ = True
                                ElseIf .indices(2) = v_index_aux Then
                                    v_index_next_aux = .indices(1)
                                    aux_must_reverse_orderQ = True
                                    match_foundQ = True
                                ElseIf .indices(3) = v_index_aux Then
                                    v_index_next_aux = .indices(0)
                                    match_foundQ = True
                                End If
                                
                                'If the vertex was found on this edge, make sure it's connected to one of the still unsorted vertices
                                If match_foundQ Then
                                    If check_opposite_directionQ Then
                                        aux_must_reverse_orderQ = Not aux_must_reverse_orderQ
                                        next_vector_list_index = GetFirstIndexOccurrenceLong(vertex_verts_out(vgi).indices, 0, _
                                                                                             num_verts_per_vert - 1 - ordered_list_index, v_index_next_aux)
                                    Else
                                        next_vector_list_index = GetFirstIndexOccurrenceLong(vertex_verts_out(vgi).indices, ordered_list_index, _
                                                                                             num_verts_per_vert - 1, v_index_next_aux)
                                    End If
                                End If
                                
                                If next_vector_list_index > -1 Then
                                    If aux_must_reverse_orderQ Then
                                        must_reverse_orderQ = True
                                    End If
                                    edge_usages(ei) = edge_usages(ei) + 1
                                    last_edge_index = ei
                                    Exit For
                                End If
                                match_foundQ = False
                            End If
                        End If
                    End With
                End If
            Next ei
            
            If Not match_foundQ Then
                If check_opposite_directionQ Then
                    Debug.Assert "OhGodWhy"
                End If
                'Couldn't find the next vertex. This means there is a gap. Move the already computed data to the end and traverse on the opposite direction.
                check_opposite_directionQ = True
                v_index_aux = vertex_verts_out(vgi).indices(0)
                v_index_next_aux = v_index_aux
                last_edge_index = -1
                For vi = 0 To ordered_list_index - 1
                    EchangeVectorElementsLong vertex_verts_out(vgi).indices, num_verts_per_vert - 1 - vi, ordered_list_index - 1 - vi
                Next vi
            Else
                If check_opposite_directionQ Then
                    EchangeVectorElementsLong vertex_verts_out(vgi).indices, num_verts_per_vert - 1 - ordered_list_index, next_vector_list_index
                Else
                    EchangeVectorElementsLong vertex_verts_out(vgi).indices, ordered_list_index, next_vector_list_index
                End If
                ordered_list_index = ordered_list_index + 1
                v_index_aux = v_index_next_aux
            End If
        Wend
        
        'If the vertices where traversed on the wrong order, invert it
        If must_reverse_orderQ Then
            With vertex_verts_out(vgi)
                For vi = 0 To (num_verts_per_vert - 1) \ 2
                  v_index_aux = .indices(vi)
                  .indices(vi) = .indices((num_verts_per_vert - 1) - vi)
                  .indices((num_verts_per_vert - 1) - vi) = v_index_aux
                Next vi
            End With
        End If
    Next vgi
End Sub
'Creates the table of vertices group at the extremes of each edge
Private Sub CreateVertGroupsPerEdgeTable(ByRef Verts() As Point3D, ByRef edge_verts() As VertsGroup, ByRef vertex_verts() As VertsGroup, _
                                         ByRef verts_per_edge_out() As Long)
    Dim ei As Long
    Dim PI As Long
    Dim vi As Long
    Dim vie As Long
    
    Dim num_edges As Long
    Dim num_verts_per_edge As Long
    Dim num_polys As Long
    Dim num_verts As Long
    Dim num_verts_per_vert As Long
    
    Dim v_index As Long
    
    num_edges = UBound(edge_verts) + 1
    num_verts = UBound(vertex_verts) + 1
    
    ReDim verts_per_edge_out(num_edges - 1, 1)
    
    For ei = 0 To num_edges - 1
        num_verts_per_edge = UBound(edge_verts(ei).indices) + 1
        For vie = 0 To 1
            For vi = 0 To num_verts - 1
                If CompareSimilarPoints3D(Verts(edge_verts(ei).indices(vie)), Verts(vertex_verts(vi).indices(0))) Then
                    verts_per_edge_out(ei, vie) = vi
                    Exit For
                End If
            Next vi
        Next vie
    Next ei
End Sub
'Computes the table of polygons adjacents to vertices
Private Sub ComputePolysPerEdgeTable(ByRef polys() As PPolygon, ByRef polys_per_edge_out() As VertsGroup)
    Dim PI As Long
    Dim ei As Long
    
    Dim num_group_polys As Long
    Dim num_polys As Long
    
    num_polys = UBound(polys) + 1
    
    For PI = 0 To num_polys - 1
        For ei = 0 To 2
            With polys_per_edge_out(polys(PI).Edges(ei))
                If SafeArrayGetDim(.indices) <> 0 Then
                    num_group_polys = UBound(.indices) + 1
                Else
                    num_group_polys = 0
                End If
                ReDim Preserve .indices(num_group_polys)
                .indices(num_group_polys) = PI
            End With
        Next ei
    Next PI
End Sub
'Contracts polygons following the Doo-Sabin smoothing rules
Private Sub DooSabinPolysContraction(ByRef polys() As PPolygon, ByRef Verts() As Point3D, ByRef v_colors() As color, ByRef tex_coords() As Point2D, _
                                     ByRef edge_verts() As VertsGroup, ByRef vertex_verts() As VertsGroup, ByRef per_vertex_coefs_out() As Point2D)
    Dim PI As Long
    Dim vi As Long
    
    Dim vi_prev As Long
    Dim vi_next As Long
    
    Dim ei_to_prev As Long
    Dim ei_to_next As Long
    Dim e_index_to_prev As Long
    Dim e_index_to_next As Long
    
    Dim poly_center As Point3D
    Dim color_center(4) As Single
    Dim tex_coord_center As Point2D
    Dim mid_point1 As Point3D
    Dim mid_point2 As Point3D
    Dim mid_point_color1 As color
    Dim mid_point_color2 As color
    Dim mid_point_tex_coord1 As Point2D
    Dim mid_point_tex_coord2 As Point2D
    Dim mid_point1_valid As Boolean
    Dim mid_point2_valid As Boolean
    
    Dim normals_angle_cos As Double
    Dim normal_current As Point3D
    Dim normal_aux As Point3D
    Dim normal_aux1 As Point3D
    Dim normal_aux2 As Point3D
    Dim adjacent_poly_index As Long
    Dim adjacent_poly_index1 As Long
    Dim adjacent_poly_index2 As Long
    
    Dim num_polys_per_edge_1 As Long
    Dim num_polys_per_edge_2 As Long
    
    Dim poly_normals() As Point3D
    Dim poly_normals_computed() As Boolean
    
    Dim aux_vert As Point3D
    Dim aux_v_color As color
    Dim aux_tex_coord As Point2D
    Dim avg_v_color(4) As Single
    Dim temp_verts(2) As Point3D
    Dim temp_v_colors(2) As color
    Dim temp_tex_coords(2) As Point2D
    Dim alpha_prev As Single
    Dim alpha_next As Single
    
    Dim num_verts As Long
    Dim num_polys  As Long
    num_verts = UBound(Verts) + 1
    num_polys = UBound(polys) + 1
    
    Dim has_tex_coordsQ As Boolean
    has_tex_coordsQ = SafeArrayGetDim(tex_coords) > 0
    
    Dim polys_per_edge() As VertsGroup
    ReDim polys_per_edge(UBound(edge_verts))
    ComputePolysPerEdgeTable polys, polys_per_edge
    
    ReDim poly_normals(num_polys)
    ReDim poly_normals_computed(num_polys)
    ZeroMemory poly_normals_computed(0), num_polys * 2
    
    ReDim per_vertex_coefs_out(num_verts - 1)
    ZeroMemory per_vertex_coefs_out(0), num_verts * 2 * 4
    
    For PI = 0 To num_polys - 1
        With polys(PI)
            If poly_normals_computed(PI) Then
                normal_current = poly_normals(PI)
            Else
                normal_current = CalculateNormal(Verts(.Verts(0)), Verts(.Verts(1)), Verts(.Verts(2)))
                normal_current = Normalize(normal_current)
                poly_normals(PI) = normal_current
                poly_normals_computed(PI) = True
            End If
            ZeroMemory poly_center, 3 * 4
            ZeroMemory color_center(0), 4 * 4
            ZeroMemory tex_coord_center, 2 * 4
            For vi = 0 To 2
                aux_vert = Verts(.Verts(vi))
                poly_center.x = poly_center.x + aux_vert.x
                poly_center.y = poly_center.y + aux_vert.y
                poly_center.z = poly_center.z + aux_vert.z
                
                aux_v_color = v_colors(.Verts(vi))
                color_center(0) = color_center(0) + CSng(aux_v_color.r)
                color_center(1) = color_center(1) + CSng(aux_v_color.g)
                color_center(2) = color_center(2) + CSng(aux_v_color.B)
                color_center(3) = color_center(3) + CSng(aux_v_color.a)
                
                If has_tex_coordsQ Then
                    aux_tex_coord = tex_coords(.Verts(vi))
                    tex_coord_center.x = tex_coord_center.x + aux_tex_coord.x
                    tex_coord_center.y = tex_coord_center.y + aux_tex_coord.y
                End If
            Next vi
            poly_center.x = poly_center.x / 3
            poly_center.y = poly_center.y / 3
            poly_center.z = poly_center.z / 3
            
            color_center(0) = color_center(0) / 3
            color_center(1) = color_center(1) / 3
            color_center(2) = color_center(2) / 3
            color_center(3) = color_center(3) / 3
            
            If has_tex_coordsQ Then
                tex_coord_center.x = tex_coord_center.x / 3
                tex_coord_center.y = tex_coord_center.y / 3
            End If
            
            For vi = 0 To 2
                aux_vert = Verts(.Verts(vi))
                aux_v_color = v_colors(.Verts(vi))
                If has_tex_coordsQ Then
                    aux_tex_coord = tex_coords(.Verts(vi))
                End If
                
                vi_next = (vi + 1) Mod 3
                vi_prev = (vi + 2) Mod 3
                
                ei_to_prev = vi
                ei_to_next = vi_prev
                
                e_index_to_prev = .Edges(ei_to_prev)
                e_index_to_next = .Edges(ei_to_next)
                
                mid_point1_valid = False
                mid_point2_valid = False
                num_polys_per_edge_1 = UBound(polys_per_edge(e_index_to_next).indices) + 1
                num_polys_per_edge_2 = UBound(polys_per_edge(e_index_to_prev).indices) + 1
                
                If num_polys_per_edge_1 = 2 And num_polys_per_edge_2 = 2 Then
                    'Both edges connect two polygons. The angle between them must be lower than the maximum.
                    adjacent_poly_index1 = polys_per_edge(e_index_to_next).indices(0)
                    If adjacent_poly_index1 = PI Then
                        adjacent_poly_index1 = polys_per_edge(e_index_to_next).indices(1)
                    End If
                    
                    adjacent_poly_index2 = polys_per_edge(e_index_to_prev).indices(0)
                    If adjacent_poly_index2 = PI Then
                        adjacent_poly_index2 = polys_per_edge(e_index_to_prev).indices(1)
                    End If
                    
                    If poly_normals_computed(adjacent_poly_index1) Then
                        normal_aux1 = poly_normals(adjacent_poly_index1)
                    Else
                        normal_aux1 = CalculateNormal(Verts(polys(adjacent_poly_index1).Verts(0)), _
                                                      Verts(polys(adjacent_poly_index1).Verts(1)), _
                                                      Verts(polys(adjacent_poly_index1).Verts(2)))
                        normal_aux1 = Normalize(normal_aux1)
                        poly_normals(adjacent_poly_index1) = normal_aux1
                        poly_normals_computed(adjacent_poly_index1) = True
                    End If
                    
                    If poly_normals_computed(adjacent_poly_index2) Then
                        normal_aux2 = poly_normals(adjacent_poly_index2)
                    Else
                        normal_aux2 = CalculateNormal(Verts(polys(adjacent_poly_index2).Verts(0)), _
                                                      Verts(polys(adjacent_poly_index2).Verts(1)), _
                                                      Verts(polys(adjacent_poly_index2).Verts(2)))
                        normal_aux2 = Normalize(normal_aux2)
                        poly_normals(adjacent_poly_index2) = normal_aux2
                        poly_normals_computed(adjacent_poly_index2) = True
                    End If
                    
                    normals_angle_cos = ComputeVectorsAngleCos(normal_aux1, normal_aux2)
                    If normals_angle_cos > MIN_SMOOTH_COS Then
                        mid_point1_valid = True
                        mid_point2_valid = True
                    End If
                Else
                    'One of the edges belongs to a hole in the mesh. To prevent mesh erosion, it's direction and position must be preserved.
                    If num_polys_per_edge_1 = 2 Then
                        'The edge 1 connects two polygons. The angle between both must be lower than the maximum.
                        adjacent_poly_index = polys_per_edge(e_index_to_next).indices(0)
                        If adjacent_poly_index = PI Then
                            adjacent_poly_index = polys_per_edge(e_index_to_next).indices(1)
                        End If
                            
                        If poly_normals_computed(adjacent_poly_index) Then
                            normal_aux = poly_normals(adjacent_poly_index)
                        Else
                            normal_aux = CalculateNormal(Verts(polys(adjacent_poly_index).Verts(0)), _
                                                         Verts(polys(adjacent_poly_index).Verts(1)), _
                                                         Verts(polys(adjacent_poly_index).Verts(2)))
                            normal_aux = Normalize(normal_aux)
                            poly_normals(adjacent_poly_index) = normal_aux
                            poly_normals_computed(adjacent_poly_index) = True
                        End If
                        
                        normals_angle_cos = ComputeVectorsAngleCos(normal_current, normal_aux)
                        mid_point1_valid = (normals_angle_cos > MIN_SMOOTH_COS)
                    End If
                    
                    If num_polys_per_edge_2 = 2 Then
                        'The edge 2 connects two polygons. The angle between both must be lower than the maximum.
                        adjacent_poly_index = polys_per_edge(e_index_to_prev).indices(0)
                        If adjacent_poly_index = PI Then
                            adjacent_poly_index = polys_per_edge(e_index_to_prev).indices(1)
                        End If
                            
                        If poly_normals_computed(adjacent_poly_index) Then
                            normal_aux = poly_normals(adjacent_poly_index)
                        Else
                            normal_aux = CalculateNormal(Verts(polys(adjacent_poly_index).Verts(0)), _
                                                         Verts(polys(adjacent_poly_index).Verts(1)), _
                                                         Verts(polys(adjacent_poly_index).Verts(2)))
                            normal_aux = Normalize(normal_aux)
                            poly_normals(adjacent_poly_index) = normal_aux
                            poly_normals_computed(adjacent_poly_index) = True
                        End If
                        
                        normals_angle_cos = ComputeVectorsAngleCos(normal_current, normal_aux)
                        
                        mid_point2_valid = (normals_angle_cos > MIN_SMOOTH_COS)
                    End If
                End If
                
                If mid_point1_valid Then
                    mid_point1 = GetPointInLine(aux_vert, Verts(.Verts(vi_next)), 0.5)
                    mid_point_color1 = InterpolateColor(v_colors(.Verts(vi)), v_colors(.Verts(vi_next)), 0.5)
                    If has_tex_coordsQ Then
                        mid_point_tex_coord1 = InterpolatePoint2D(tex_coords(.Verts(vi)), tex_coords(.Verts(vi_next)), 0.5)
                    End If
                End If
                
                If mid_point2_valid Then
                    mid_point2 = GetPointInLine(aux_vert, Verts(.Verts(vi_prev)), 0.5)
                    mid_point_color2 = InterpolateColor(v_colors(.Verts(vi)), v_colors(.Verts(vi_prev)), 0.5)
                    If has_tex_coordsQ Then
                        mid_point_tex_coord2 = InterpolatePoint2D(tex_coords(.Verts(vi)), tex_coords(.Verts(vi_prev)), 0.5)
                    End If
                End If
                
                If mid_point1_valid And mid_point2_valid Then
                    temp_verts(vi) = poly_center
                    
                    temp_verts(vi).x = temp_verts(vi).x + mid_point1.x
                    temp_verts(vi).y = temp_verts(vi).y + mid_point1.y
                    temp_verts(vi).z = temp_verts(vi).z + mid_point1.z
                    
                    temp_verts(vi).x = temp_verts(vi).x + mid_point2.x
                    temp_verts(vi).y = temp_verts(vi).y + mid_point2.y
                    temp_verts(vi).z = temp_verts(vi).z + mid_point2.z
                    
                    temp_verts(vi).x = temp_verts(vi).x + aux_vert.x
                    temp_verts(vi).y = temp_verts(vi).y + aux_vert.y
                    temp_verts(vi).z = temp_verts(vi).z + aux_vert.z
                    
                    temp_verts(vi).x = temp_verts(vi).x / 4
                    temp_verts(vi).y = temp_verts(vi).y / 4
                    temp_verts(vi).z = temp_verts(vi).z / 4
                    
                    
                    CopyMemory avg_v_color(0), color_center(0), 4 * 4
                    
                    avg_v_color(0) = avg_v_color(0) + CSng(mid_point_color1.r)
                    avg_v_color(1) = avg_v_color(1) + CSng(mid_point_color1.g)
                    avg_v_color(2) = avg_v_color(2) + CSng(mid_point_color1.B)
                    avg_v_color(3) = avg_v_color(3) + CSng(mid_point_color1.a)
                    
                    avg_v_color(0) = avg_v_color(0) + CSng(mid_point_color2.r)
                    avg_v_color(1) = avg_v_color(1) + CSng(mid_point_color2.g)
                    avg_v_color(2) = avg_v_color(2) + CSng(mid_point_color2.B)
                    avg_v_color(3) = avg_v_color(3) + CSng(mid_point_color2.a)
                    
                    avg_v_color(0) = avg_v_color(0) + CSng(aux_v_color.r)
                    avg_v_color(1) = avg_v_color(1) + CSng(aux_v_color.g)
                    avg_v_color(2) = avg_v_color(2) + CSng(aux_v_color.B)
                    avg_v_color(3) = avg_v_color(3) + CSng(aux_v_color.a)
                    
                    temp_v_colors(vi).r = CByte(Min(255, avg_v_color(0) / 4))
                    temp_v_colors(vi).g = CByte(Min(255, avg_v_color(1) / 4))
                    temp_v_colors(vi).B = CByte(Min(255, avg_v_color(2) / 4))
                    temp_v_colors(vi).a = CByte(Min(255, avg_v_color(3) / 4))
                    
                    
                    If has_tex_coordsQ Then
                        temp_tex_coords(vi) = tex_coord_center
                        
                        temp_tex_coords(vi).x = temp_tex_coords(vi).x + mid_point_tex_coord1.x
                        temp_tex_coords(vi).y = temp_tex_coords(vi).y + mid_point_tex_coord1.y
                        
                        temp_tex_coords(vi).x = temp_tex_coords(vi).x + mid_point_tex_coord2.x
                        temp_tex_coords(vi).y = temp_tex_coords(vi).y + mid_point_tex_coord2.y
                        
                        temp_tex_coords(vi).x = temp_tex_coords(vi).x + aux_tex_coord.x
                        temp_tex_coords(vi).y = temp_tex_coords(vi).y + aux_tex_coord.y
                        
                        temp_tex_coords(vi).x = temp_tex_coords(vi).x / 4
                        temp_tex_coords(vi).y = temp_tex_coords(vi).y / 4
                    End If
                    
                    
                    per_vertex_coefs_out(.Verts(vi)).x = CalculatePoint2LineProjectionPosition(temp_verts(vi), Verts(.Verts(vi_next)), aux_vert)
                    per_vertex_coefs_out(.Verts(vi)).y = CalculatePoint2LineProjectionPosition(temp_verts(vi), Verts(.Verts(vi_prev)), aux_vert)
                ElseIf mid_point1_valid Then
                    per_vertex_coefs_out(.Verts(vi)).x = 0.75
                    temp_verts(vi) = GetPointInLine(aux_vert, Verts(.Verts(vi_next)), 0.25)
                    temp_v_colors(vi) = InterpolateColor(v_colors(.Verts(vi)), v_colors(.Verts(vi_next)), 0.25)
                    If has_tex_coordsQ Then
                        temp_tex_coords(vi) = InterpolatePoint2D(tex_coords(.Verts(vi)), tex_coords(.Verts(vi_next)), 0.25)
                    End If
                ElseIf mid_point2_valid Then
                    per_vertex_coefs_out(.Verts(vi)).y = 0.75
                    temp_verts(vi) = GetPointInLine(aux_vert, Verts(.Verts(vi_prev)), 0.25)
                    temp_v_colors(vi) = InterpolateColor(v_colors(.Verts(vi)), v_colors(.Verts(vi_prev)), 0.25)
                    If has_tex_coordsQ Then
                        temp_tex_coords(vi) = InterpolatePoint2D(tex_coords(.Verts(vi)), tex_coords(.Verts(vi_prev)), 0.25)
                    End If
                Else
                    temp_verts(vi) = aux_vert
                    temp_v_colors(vi) = aux_v_color
                    temp_tex_coords(vi) = aux_tex_coord
                End If
            Next vi
            
            For vi = 0 To 2
                CopyMemory Verts(.Verts(vi)), temp_verts(vi), 3 * 4
                CopyMemory v_colors(.Verts(vi)), temp_v_colors(vi), 4
                If has_tex_coordsQ Then
                    CopyMemory tex_coords(.Verts(vi)), temp_tex_coords(vi), 2 * 4
                End If
            Next vi
        End With
    Next PI
End Sub
'Adds a new triangle and returns it's index
Private Function AppendTriangle(ByRef polys() As PPolygon, ByVal v_index1 As Integer, ByVal v_index2 As Integer, ByVal v_index3 As Integer) As Long
    Dim num_polys As Long
    
    num_polys = UBound(polys) + 1
    ReDim Preserve polys(num_polys)
    With polys(num_polys)
        .Verts(0) = v_index1
        .Verts(1) = v_index2
        .Verts(2) = v_index3
        
        .Edges(0) = -1
        .Edges(1) = -1
        .Edges(2) = -1
    End With
    
    AppendTriangle = num_polys
End Function
'Adds a pair of triangle (forming a quad) and returns the index of the first one (the second one is correlative)
Private Function AppendQuad(ByRef polys() As PPolygon, ByRef verts_group As VertsGroup) As Long
    With verts_group
        If UBound(.indices) = 3 Then
            AppendQuad = AppendTriangle(polys, .indices(2), .indices(1), .indices(0))
            AppendTriangle polys, .indices(3), .indices(2), .indices(0)
        End If
    End With
End Function

'Connects the vertices at where at the same edge befor calling DooSabinPolysContraction. Returns the list of triangles added per edge.
Private Sub ConnectEdgeVertices(ByRef polys() As PPolygon, ByRef Verts() As Point3D, ByRef v_colors() As color, ByRef tex_coords() As Point2D, _
                                ByRef edge_verts() As VertsGroup, ByRef edge_polys_out() As VertsGroup)
    Dim ei As Long
    Dim pi_new As Long
    
    Dim num_edges As Long
    num_edges = UBound(edge_verts) + 1
    
    ReDim edge_polys_out(num_edges - 1)
    For ei = 0 To num_edges - 1
        pi_new = AppendQuad(polys, edge_verts(ei))
        With edge_polys_out(ei)
            ReDim .indices(1)
            .indices(0) = pi_new
            .indices(1) = pi_new + 1
        End With
    Next ei
End Sub
Private Function AppendTrinagleFan(ByRef polys() As PPolygon, ByRef external_verts() As Long, ByVal v_center_index As Integer) As Long
    Dim vi As Long
    Dim num_external_verts As Long
    Dim last_pi_new As Long
    
    num_external_verts = UBound(external_verts) + 1
    For vi = 0 To num_external_verts - 2
        AppendTriangle polys, v_center_index, external_verts(vi), external_verts(vi + 1)
    Next vi
    last_pi_new = AppendTriangle(polys, v_center_index, external_verts(num_external_verts - 1), external_verts(0))
    
    AppendTrinagleFan = last_pi_new - num_external_verts - 1
End Function
'Connects the vertices that where colapsed befor calling DooSabinPolysContraction. The new polygon will be triangulated as a fan (adding a vertex at the center)
Private Sub ConnectVertexVertices(ByRef polys() As PPolygon, ByRef Verts() As Point3D, ByRef v_colors() As color, ByRef tex_coords() As Point2D, _
                                  ByRef v_colors_original() As color, ByRef tex_coords_original() As Point2D, ByRef vertex_verts() As VertsGroup, ByRef vertex_verts_center_out() As Long, ByRef vertex_polys_out() As VertsGroup)
    Dim vgi As Long
    Dim vi As Long
    Dim PI As Long
    Dim v_index_aux As Long
    Dim center_vertex_index As Long
    Dim v_color_aux As color
    Dim tex_coord_aux As Point2D
    
    Dim aux_d As Single
    Dim fan_center As Point3D
    Dim fan_center_a As Point3D
    Dim fan_center_b As Point3D
    Dim aux_length_a As Single
    Dim aux_length_b As Single
    Dim longuest_edge_a As Single
    Dim longuest_edge_b As Single
    
    Dim first_fan_triangle_index As Long
    
    Dim num_verts As Long
    Dim num_group_verts As Long
    num_group_verts = UBound(vertex_verts) + 1
    
    Dim has_tex_coordsQ As Boolean
    has_tex_coordsQ = SafeArrayGetDim(tex_coords) > 0
    
    Dim dummy_normal As Point3D
    Dim dummy_normals() As Point3D
    
    ReDim vertex_polys_out(num_group_verts - 1)
    ReDim vertex_verts_center_out(num_group_verts - 1)
    For vgi = 0 To num_group_verts - 1
        'Polygon center A: Compute vertices average
        ZeroMemory fan_center_a, 3 * 4
        num_verts = UBound(vertex_verts(vgi).indices) + 1
        With fan_center_a
            For vi = 0 To num_verts - 1
                v_index_aux = vertex_verts(vgi).indices(vi)
                .x = .x + Verts(v_index_aux).x
                .y = .y + Verts(v_index_aux).y
                .z = .z + Verts(v_index_aux).z
            Next vi
            
            .x = .x / num_verts
            .y = .y / num_verts
            .z = .z / num_verts
        End With
        
        'Polygon center B: Compute orthogonal projection
        With vertex_verts(vgi)
            aux_d = ComputePlaneD(.normal, Verts(.indices(0)))
            fan_center_b = GetPoint3DOrthogonalProjection(.OriginalPosition, .normal.x, .normal.y, .normal.z, aux_d)
        End With
        
        'Pick the vertex which creates shortest edges
        longuest_edge_a = 0
        longuest_edge_b = 0
        For vi = 0 To num_verts - 1
            v_index_aux = vertex_verts(vgi).indices(vi)
            aux_length_a = CalculateDistance(fan_center_a, Verts(v_index_aux))
            aux_length_b = CalculateDistance(fan_center_b, Verts(v_index_aux))
            longuest_edge_a = IIf(aux_length_a > longuest_edge_a, aux_length_a, longuest_edge_a)
            longuest_edge_b = IIf(aux_length_b > longuest_edge_b, aux_length_b, longuest_edge_b)
        Next vi
        
        If longuest_edge_a > longuest_edge_b Then
            fan_center = fan_center_b
        Else
            fan_center = fan_center_a
        End If
        
        v_index_aux = vertex_verts(vgi).indices(0)
        CopyMemory v_color_aux, v_colors_original(v_index_aux), 4
        If has_tex_coordsQ Then
            CopyMemory tex_coord_aux, tex_coords_original(v_index_aux), 2 * 4
            center_vertex_index = AppendVertexWithTexCoords(fan_center, dummy_normal, v_color_aux, tex_coord_aux, Verts, dummy_normals, v_colors, tex_coords)
        Else
            center_vertex_index = AppendVertex(fan_center, dummy_normal, v_color_aux, Verts, dummy_normals, v_colors)
        End If
        vertex_verts_center_out(vgi) = center_vertex_index
        first_fan_triangle_index = AppendTrinagleFan(polys, vertex_verts(vgi).indices, center_vertex_index)
        ReDim vertex_polys_out(vgi).indices(num_verts - 1)
        For PI = 0 To num_verts - 1
            vertex_polys_out(vgi).indices(PI) = first_fan_triangle_index + PI
        Next PI
    Next vgi
End Sub

'Splits vertex and edge polys to match the projection of the original geometry bounds. Also fixes vertex atributes (vertex color and tex coords).
Private Sub FixVertexAtributes(ByRef polys() As PPolygon, ByRef Verts() As Point3D, ByRef Normals() As Point3D, ByRef v_colors() As color, ByRef tex_coords() As Point2D, _
                               ByRef original_v_colors() As color, ByRef original_tex_coords() As Point2D, ByRef per_vertex_coefs() As Point2D, _
                               ByRef per_edge_verts() As VertsGroup, ByRef edge_polys() As VertsGroup, ByRef vertex_verts_center() As Long, _
                               ByRef vertex_polys() As VertsGroup, ByRef vert_groups_per_edge() As Long)
    Dim ei As Long
    Dim PI As Long
    Dim vi As Long
    
    Dim v1_index As Long
    Dim v2_index As Long
    
    Dim v_edge0 As Point3D
    Dim v_edge1 As Point3D
    Dim v_edge2 As Point3D
    Dim v_edge3 As Point3D
    
    Dim vect1 As Point3D
    Dim vect2 As Point3D
    Dim a As Single
    Dim B As Single
    Dim C As Single
    Dim d As Single
    
    Dim cut1_validQ As Boolean
    Dim cut2_validQ As Boolean
    Dim cut_alpha1 As Double
    Dim cut_alpha2 As Double
    Dim cut_point1 As Point3D
    Dim cut_point2 As Point3D
    Dim copy_attribs(5) As Long
    Dim aux_v1 As Point3D
    Dim aux_v2 As Point3D
    Dim aux_v_color(3) As color
    Dim aux_tex_coord(3) As Point2D
    
    Dim vertex_indices_copy(3) As Long
    Dim v_colors_copy(3) As color
    Dim tex_coords_copy(3) As Point2D
    Dim new_verts_indices(3) As Long
    
    Dim dummy As Point3D
    dummy.x = 0
    dummy.y = 0
    dummy.z = 0
    
    Dim num_edges As Long
    Dim num_polys As Long
    
    num_edges = UBound(edge_polys) + 1
    
    Dim has_tex_coordsQ As Boolean
    has_tex_coordsQ = SafeArrayGetDim(tex_coords) > 0
    
    Dim must_separate_attributesQ As Boolean
    
    For ei = 0 To num_edges - 1
        must_separate_attributesQ = False
        If SafeArrayGetDim(per_edge_verts(ei).indices) > 0 Then
            With per_edge_verts(ei)
                must_separate_attributesQ = Not CompareColors(original_v_colors(.indices(0)), original_v_colors(.indices(3))) Or _
                                            Not CompareColors(original_v_colors(.indices(1)), original_v_colors(.indices(2)))
                If Not must_separate_attributesQ And has_tex_coordsQ Then
                    must_separate_attributesQ = Not ComparePoints2D(original_tex_coords(.indices(0)), original_tex_coords(.indices(3))) Or _
                                                Not ComparePoints2D(original_tex_coords(.indices(1)), original_tex_coords(.indices(2)))
                End If
            End With
        End If
        
        'There is no need to cut the geometry if there is no discontinuty in the attributes.
        If must_separate_attributesQ Then
            'Separate edge polygons for ease sake
            With per_edge_verts(ei)
                For vi = 0 To 3
                    vertex_indices_copy(vi) = .indices(vi)
                    If has_tex_coordsQ Then
                        .indices(vi) = AppendVertexWithTexCoords(Verts(.indices(vi)), dummy, v_colors(.indices(vi)), tex_coords(.indices(vi)), _
                                                                 Verts, Normals, v_colors, tex_coords)
                    Else
                        .indices(vi) = AppendVertex(Verts(.indices(vi)), dummy, v_colors(.indices(vi)), Verts, Normals, v_colors)
                    End If
                Next vi
            End With
            
            ReDim Preserve per_edge_verts(ei).indices(5)
            With per_edge_verts(ei)
                If has_tex_coordsQ Then
                    .indices(4) = AppendVertexWithTexCoords(Verts(.indices(0)), dummy, v_colors(.indices(0)), tex_coords(.indices(0)), _
                                                            Verts, Normals, v_colors, tex_coords)
                    .indices(5) = AppendVertexWithTexCoords(Verts(.indices(2)), dummy, v_colors(.indices(2)), tex_coords(.indices(2)), _
                                                            Verts, Normals, v_colors, tex_coords)
                    
                Else
                    .indices(4) = AppendVertex(Verts(.indices(0)), dummy, v_colors(.indices(0)), Verts, Normals, v_colors)
                    .indices(5) = AppendVertex(Verts(.indices(2)), dummy, v_colors(.indices(2)), Verts, Normals, v_colors)
                End If
            End With
            
            With polys(edge_polys(ei).indices(0))
                .Verts(0) = per_edge_verts(ei).indices(2)
                .Verts(1) = per_edge_verts(ei).indices(1)
                .Verts(2) = per_edge_verts(ei).indices(0)
            End With
            
            With polys(edge_polys(ei).indices(1))
                .Verts(0) = per_edge_verts(ei).indices(3)
                .Verts(1) = per_edge_verts(ei).indices(5)
                .Verts(2) = per_edge_verts(ei).indices(4)
            End With
            
            'Compute cut plane
            v1_index = vert_groups_per_edge(ei, 0)
            v2_index = vert_groups_per_edge(ei, 1)
            vect1 = Verts(v1_index)
            vect2 = Verts(v2_index)
            
            With Verts(vertex_verts_center(v1_index)) '.OriginalPosition
                vect1.x = vect1.x - .x
                vect1.y = vect1.y - .y
                vect1.z = vect1.z - .z
            End With
            
            With Verts(vertex_verts_center(v2_index)) '.OriginalPosition
                vect2.x = vect2.x - .x
                vect2.y = vect2.y - .y
                vect2.z = vect2.z - .z
            End With
            
            ComputePlaneABCD vect1, vect2, Verts(v1_index), a, B, C, d
            
            With per_edge_verts(ei)
                v_edge0 = Verts(.indices(0))
                v_edge1 = Verts(.indices(1))
                v_edge2 = Verts(.indices(2))
                v_edge3 = Verts(.indices(3))
            End With
            'Compute cut points
            cut1_validQ = GetVectorToPlaneIntersection(v_edge0, v_edge3, a, B, C, d, cut_alpha1)
            cut2_validQ = GetVectorToPlaneIntersection(v_edge1, v_edge2, a, B, C, d, cut_alpha2)
            
            copy_attribs(0) = -1
            copy_attribs(1) = -1
            copy_attribs(2) = -1
            copy_attribs(3) = -1
            
            'Decide where attirbutes must be copied to and handle extreme cases
            If cut1_validQ Then
                If cut_alpha1 < 0.0001 Then
                    cut_point1 = v_edge0
                    copy_attribs(0) = 3
                    cut1_validQ = False
                ElseIf cut_alpha1 > 0.9999 Then
                    cut_point1 = v_edge3
                    copy_attribs(3) = 0
                    cut1_validQ = False
                Else
                    cut_point1 = GetPointInLine(v_edge0, v_edge3, cut_alpha1)
                    copy_attribs(3) = 0
                End If
            Else
                cut_point1 = v_edge0
                copy_attribs(0) = 3
            End If
            
            If cut2_validQ Then
                If cut_alpha2 < 0.0001 Then
                    cut_point2 = v_edge1
                    copy_attribs(1) = 2
                    cut2_validQ = False
                ElseIf cut_alpha2 > 0.9999 Then
                    cut_point2 = v_edge2
                    copy_attribs(2) = 1
                    cut2_validQ = False
                Else
                    cut_point2 = GetPointInLine(v_edge1, v_edge2, cut_alpha1)
                    copy_attribs(2) = 1
                End If
            Else
                cut_point2 = v_edge1
                copy_attribs(1) = 2
            End If
            
            copy_attribs(4) = copy_attribs(0)
            copy_attribs(5) = copy_attribs(2)
            
            With per_edge_verts(ei)
                aux_v1 = Verts(.indices(3))
                aux_v2 = Verts(.indices(2))
                'Move the atributes to the border of the edges
                For vi = 0 To 5
                    If copy_attribs(vi) > -1 Then
                        v_colors(.indices(vi)) = v_colors(.indices(copy_attribs(vi)))
                        If has_tex_coordsQ Then
                            tex_coords(.indices(vi)) = tex_coords(.indices(copy_attribs(vi)))
                        End If
                    End If
                Next vi
            
                'Cut the edges of the edge quad
                If cut1_validQ Or cut2_validQ Then
                    If cut1_validQ Then
                        Verts(.indices(3)) = cut_point1
                    End If
                    
                    If cut2_validQ Then
                        Verts(.indices(2)) = cut_point2
                        Verts(.indices(5)) = cut_point2
                    End If
                
                    If cut1_validQ And cut2_validQ Then
                        '.y -> next, .x -> prev
                        'aux_v_color(0) = InterpolateColor(original_v_colors(vertex_indices_copy(0)), original_v_colors(vertex_indices_copy(1)), _
                        '                                  per_vertex_coefs(vertex_indices_copy(0).y))
                        'aux_v_color(1) = InterpolateColor(original_v_colors(vertex_indices_copy(1)), original_v_colors(vertex_indices_copy(0)), _
                        '                                  per_vertex_coefs(vertex_indices_copy(1).x))
                        'aux_v_color(2) = original_v_colors(vertex_indices_copy(2))
                        'aux_v_color(3) = original_v_colors(vertex_indices_copy(3))
                        'If has_tex_coordsQ Then
                        '    aux_tex_coord(0) = InterpolateColor(original_tex_coords(vertex_indices_copy(0)), original_tex_coords(vertex_indices_copy(1)), _
                        '                                        per_vertex_coefs(vertex_indices_copy(0).y))
                        '    aux_tex_coord(1) = InterpolateColor(original_tex_coords(vertex_indices_copy(1)), original_tex_coords(vertex_indices_copy(0)), _
                        '                                        per_vertex_coefs(vertex_indices_copy(1).x))
                        '    aux_tex_coord(2) = original_tex_coords(vertex_indices_copy(2))
                        '    aux_tex_coord(3) = original_tex_coords(vertex_indices_copy(3))
                        
                        '    new_verts(0) = AppendVertexWithTexCoords(cut_point1, dummy, aux_v_color(0), aux_tex_coord(0), Verts, Normals, v_colors, tex_coords)
                        '    new_verts(1) = AppendVertexWithTexCoords(cut_point2, dummy, aux_v_color(1), aux_tex_coord(1), Verts, Normals, v_colors, tex_coords)
                        '    new_verts(2) = AppendVertexWithTexCoords(aux_v1, dummy, aux_v_color(2), aux_tex_coord(2), Verts, Normals, v_colors, tex_coords)
                        '    new_verts(3) = AppendVertexWithTexCoords(aux_v2, dummy, aux_v_color(3), aux_tex_coord(3), Verts, Normals, v_colors, tex_coords)
                        'Else
                        '    new_verts(0) = AppendVertex(cut_point1, dummy, aux_v_color(0), Verts, Normals, v_colors)
                        '    new_verts(1) = AppendVertex(cut_point2, dummy, aux_v_color(1), Verts, Normals, v_colors)
                        '    new_verts(2) = AppendVertex(aux_v1, dummy, aux_v_color(2), Verts, Normals, v_colors)
                        '    new_verts(3) = AppendVertex(aux_v2, dummy, aux_v_color(3), Verts, Normals, v_colors)
                        'End
                        'pi_new = AppendQuad(polys, new_verts)
                        
                        
                    Else
                        
                    End If
                End If
            End With
        End If
    Next ei
End Sub
'Appends isolated data into a PModel's structures and returns the new group
Private Sub AppendIsolatedGroup(ByRef group_in As PGroup, _
                                ByRef polys() As PPolygon, ByRef Verts() As Point3D, ByRef v_colors() As color, ByRef tex_coords() As Point2D, _
                                ByRef group_out As PGroup, _
                                ByRef polys_out() As PPolygon, ByRef verts_out() As Point3D, ByRef v_colors_out() As color, ByRef tex_coords_out() As Point2D)
    Dim PI As Long
    Dim vi As Long
    Dim vertex_index As Long
    
    CopyMemory group_out, group_in, 14 * 4

    With group_out
        .DListNum = 0
        .HiddenQ = False
        
        .numPoly = UBound(polys) + 1
        .numvert = UBound(Verts) + 1
        
        If SafeArrayGetDim(polys_out) > 0 Then
            .offpoly = UBound(polys_out) + 1
        Else
            .offpoly = 0
        End If
        
        If SafeArrayGetDim(verts_out) > 0 Then
            .offvert = UBound(verts_out) + 1
        Else
            .offvert = 0
        End If
        
        If SafeArrayGetDim(tex_coords_out) > 0 Then
            .offTex = UBound(tex_coords_out) + 1
        Else
            .offTex = 0
        End If
        
        'Debug.Print ".offpoly = "; .offpoly; ", .numPoly = "; .numPoly
        'Debug.Print "Hola2: .offpoly = "; .offpoly; ", .offvert = "; .offvert
        'Debug.Print "Hola2: .numPoly = "; .numPoly; ", .numvert = "; .numvert
        'Debug.Print ""
        ReDim Preserve polys_out(.offpoly + .numPoly - 1)
        ReDim Preserve verts_out(.offvert + .numvert - 1)
        ReDim Preserve v_colors_out(.offvert + .numvert - 1)
        If .texFlag = 1 Then
            ReDim Preserve tex_coords_out(.offTex + .numvert - 1)
        End If
        
        CopyMemory polys_out(.offpoly), polys(0), .numPoly * 24
        CopyMemory verts_out(.offvert), Verts(0), .numvert * 3 * 4
        CopyMemory v_colors_out(.offvert), v_colors(0), .numvert * 4
        If .texFlag = 1 Then
            CopyMemory tex_coords_out(.offTex), tex_coords(0), .numvert * 2 * 4
        End If
        
        'For PI = .offpoly To .offpoly + .numPoly - 1
        '    For vi = 0 To 2
                'polys_out(PI).Verts(vi) = polys_out(PI).Verts(vi) + .offvert
        '        vertex_index = .offvert + polys_out(PI).Verts(vi)
        '        Debug.Print verts_out(vertex_index).x; ", "; verts_out(vertex_index).y; ", "; verts_out(vertex_index).z
        '    Next vi
        'Next PI
    End With
End Sub

