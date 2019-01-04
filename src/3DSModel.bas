Attribute VB_Name = "Model3DS_Module"
'Code ported from P.P.A.Narayanan c++ 3ds loader
'http://www.gamedev.net/reference/articles/article1259.asp

Option Explicit

Type rgb_3ds
    red As Byte
    green As Byte
    blue As Byte
End Type

Type vert_3ds
    x As Single
    z As Single 'z and y are swaped
    y As Single
End Type

Type tex_coord_3ds
    U As Single
    v As Single
End Type

Type face_3ds
    vertA As Integer
    vertB As Integer
    vertC As Integer
    flags As Integer    'From lib3ds (http://www.koders.com/cpp/fid8EDB781A3412B5309868BD6B71F9A9FE01469EDD.aspx?s=bump+map)
                        'Bit 0: Edge visibility AC
                        'Bit 1: Edge visibility BC
                        'Bit 2: Edge visibility AB
                        'Bit 3: Face is at tex U wrap seam
                        'Bit 4: Face is at tex V wrap seam
                        'Bit 5-8: Unused ?
                        'Bit 9-10: Random ?
                        'Bit 11-12: Unused ?
                        'Bit 13: Selection of the face in selection 3
                        'Bit 14: Selection of the face in selection 2
                        'Bit 15: Selection of the face in selection 1

End Type

Type map_list_node
    TextureFileName(255) As Byte 'Mapping filename (Texture)
    U As Single           'U scale
    v As Single           'V scale
    UOff As Single        'U Offset
    VOff As Single        'V Offset
    Rotation As Single    'Rotation angle
End Type

Type mat_list_node
    MaterialName(199) As Byte 'Material name
    Ambient As rgb_3ds       'Ambient color
    Diffuse As rgb_3ds       'Diffuse color
    Specular As rgb_3ds      'Specular color
    TextureMapsV() As map_list_node 'Texture maps
End Type

Type face_mat_node
    MaterialName(199) As Byte   'Material Name
    NumEntries As Integer       'No. of entries
    facesV() As Integer         'Faces assigned to this material
End Type

Type mesh_object_node
    MeshName(199) As Byte       'Object name
    vertsV() As vert_3ds        'Vertex list
    facesV() As face_3ds        'Face list
    NumVerts As Integer         'No. of vertices
    NumFaces As Integer         'No. of faces
    NumMappedVerts As Integer   'No. of vertices having mapping coords.
    TexCoordsV() As tex_coord_3ds       'Mapping coords. as U,V pairs (actual texture coordinates)
    LocalTransformationMatrix(11) As Single    'Local transformation matrix (last row is allways 0 0 0 1)
    FaceMaterialsV() As face_mat_node
    FaceMaterialIndicesV() As Integer  'Index of material for every face
End Type

Type Model3Ds
    modelName As String
    TranslationMatrix(11) As Single 'Translation matrix for objects (last row is allways 0 0 0 1)
    MeshesV() As mesh_object_node
    MaterialsV() As mat_list_node
End Type

Sub ReadMaterial3Ds(ByVal fileNumber As Integer, ByRef offset As Long, ByVal fileLength As Long, ByVal Length As Long, ByRef MaterialsV() As mat_list_node)
    Dim ci As Integer
    Dim count As Long
    Dim id As Integer
    Dim llen As Long
    Dim doneQ As Integer
    Dim isAmbientQ As Boolean
    Dim isDiffuseQ As Boolean
    Dim isSpecularQ As Boolean
    Dim material_index As Integer
    Dim map_index As Integer
    Dim red As Byte
    Dim green As Byte
    Dim blue As Byte
    
    isAmbientQ = False
    isDiffuseQ = False
    isSpecularQ = False
    doneQ = False
    count = offset + (Length - 6)

    If SafeArrayGetDim(MaterialsV) <> 0 Then
        material_index = UBound(MaterialsV) + 1
    Else
        material_index = 0
    End If
    
    ReDim Preserve MaterialsV(material_index)
    With MaterialsV(material_index)
        Do
            Get fileNumber, offset, id
            offset = offset + 2
            
            If offset >= fileLength Then ' OOPS! EOF
                doneQ = True
                Exit Do
            End If
              
            Get fileNumber, offset, llen
            offset = offset + 4
            Select Case id
                Case &HA000:
                    ' Read material name
                    ci = 0
                    Do
                        Get fileNumber, offset, .MaterialName(ci)
                        offset = offset + 1
                        ci = ci + 1
                    Loop Until .MaterialName(ci - 1) = 0
                Case &HA010:
                    'Hey! AMBIENT
                    isDiffuseQ = False
                    isSpecularQ = False
                    isAmbientQ = True
                    .Ambient.red = 0
                    .Ambient.green = 0
                    .Ambient.blue = 0
                Case &HA020:
                    'Hey! DIFFUSE
                    isDiffuseQ = True
                    isSpecularQ = False
                    isAmbientQ = False
                    .Diffuse.red = 0
                    .Diffuse.green = 0
                    .Diffuse.blue = 0
                Case &HA030:
                    'OH! SPECULAR
                    isDiffuseQ = False
                    isSpecularQ = True
                    isAmbientQ = False
                    .Specular.red = 0
                    .Specular.green = 0
                    .Specular.blue = 0
                Case &HA200:
                    'Texture
                    If SafeArrayGetDim(.TextureMapsV) <> 0 Then
                        map_index = UBound(.TextureMapsV) + 1
                    Else
                        map_index = 0
                    End If
                    ReDim Preserve .TextureMapsV(map_index)
                    .TextureMapsV(map_index).U = 0
                    .TextureMapsV(map_index).v = 0
                    .TextureMapsV(map_index).UOff = 0
                    .TextureMapsV(map_index).VOff = 0
                    .TextureMapsV(map_index).Rotation = 0
                Case &HA300:
                    'Texture name (filename with out path)
                    map_index = UBound(.TextureMapsV)
                    ci = 0
                    Do
                        Get fileNumber, offset, .TextureMapsV(map_index).TextureFileName(ci)
                        offset = offset + 1
                        ci = ci + 1
                    Loop Until .TextureMapsV(map_index).TextureFileName(ci - 1) = 0
                Case &HA354:
                    'V coords
                    map_index = UBound(.TextureMapsV)
                    Get fileNumber, offset, .TextureMapsV(map_index).v
                    offset = offset + 4
                Case &HA356:
                    'U coords
                    map_index = UBound(.TextureMapsV)
                    Get fileNumber, offset, .TextureMapsV(map_index).U
                    offset = offset + 4
                Case &HA358:
                    'U offset
                    map_index = UBound(.TextureMapsV)
                    Get fileNumber, offset, .TextureMapsV(map_index).UOff
                    offset = offset + 4
                Case &HA35A:
                    'V offset
                    map_index = UBound(.TextureMapsV)
                    Get fileNumber, offset, .TextureMapsV(map_index).VOff
                    offset = offset + 4
                Case &HA35C:
                    'Texture rotation angle
                    map_index = UBound(.TextureMapsV)
                    Get fileNumber, offset, .TextureMapsV(map_index).Rotation
                    offset = offset + 4
                Case &H11:
                    'Read colors
                    If isDiffuseQ Then
                        Get fileNumber, offset, .Diffuse
                    ElseIf isAmbientQ Then
                        Get fileNumber, offset, .Ambient
                    Else
                        Get fileNumber, offset, .Specular
                    End If
                    offset = offset + 3
                Case Else:
                    'Unknown chunk
                    If offset - 6 >= count Then
                        offset = offset - 6
                        doneQ = True
                    Else
                        offset = offset + llen - 6
                        doneQ = offset >= fileLength
                    End If
              End Select
        Loop Until doneQ
    End With
End Sub

Sub ReadMesh3Ds(ByVal fileNumber As Integer, ByRef offset As Long, ByVal fileLength As Long, ByVal Length As Long, ByRef MeshesV() As mesh_object_node)
    Dim ci As Long
    Dim count As Long
    Dim id As Integer
    Dim llen As Long
    Dim doneQ As Boolean
    Dim mesh_index As Integer
    Dim mat_index As Integer
    Dim num_materials As Integer
    Dim num_faces As Integer
    Dim foundQ As Boolean
    Dim temp_mesh As mesh_object_node
    
    Dim test_str As String
    
    count = offset + Length - 6
    
    doneQ = False
    
    With temp_mesh
        ci = 0
        Do
            Get fileNumber, offset, .MeshName(ci)
            test_str = test_str + Chr(.MeshName(ci))
            offset = offset + 1
            ci = ci + 1
        Loop Until .MeshName(ci - 1) = 0
        
        Do
            Get fileNumber, offset, id
            offset = offset + 2
            
            If (offset >= fileLength) Then
                doneQ = True
                Exit Do
            End If
            
            Get fileNumber, offset, llen
            offset = offset + 4
            
            Select Case id
                Case &H4100:
                    'Errr... don't know. Do nothing.
                Case &H4110:
                    'Read vertices chunk
                    Get fileNumber, offset, .NumVerts
                    offset = offset + 2
                    ReDim .vertsV(.NumVerts - 1)
                    Get fileNumber, offset, .vertsV
                    offset = offset + 3 * 4 * CLng(.NumVerts)
                Case &H4120:
                    'Read faces chunk
                    Get fileNumber, offset, .NumFaces
                    offset = offset + 2
                    ReDim .facesV(.NumFaces - 1)
                    Get fileNumber, offset, .facesV
                    offset = offset + 4 * 2 * CLng(.NumFaces)
                Case &H4130:
                    'Read material mapping info
                    If SafeArrayGetDim(.FaceMaterialsV) <> 0 Then
                        mat_index = UBound(.FaceMaterialsV) + 1
                    Else
                        mat_index = 0
                    End If
                    
                    ReDim Preserve .FaceMaterialsV(mat_index)
                    ci = 0
                    Do
                        Get fileNumber, offset, .FaceMaterialsV(mat_index).MaterialName(ci)
                        offset = offset + 1
                        ci = ci + 1
                    Loop Until .FaceMaterialsV(mat_index).MaterialName(ci - 1) = 0
                    
                    Get fileNumber, offset, .FaceMaterialsV(mat_index).NumEntries
                    offset = offset + 2
                    ReDim .FaceMaterialsV(mat_index).facesV(.FaceMaterialsV(mat_index).NumEntries - 1)
                    Get fileNumber, offset, .FaceMaterialsV(mat_index).facesV
                    offset = offset + 2 * .FaceMaterialsV(mat_index).NumEntries
                Case &H4140:
                    'Read texture coordinates
                    Get fileNumber, offset, .NumMappedVerts
                    offset = offset + 2
                    ReDim .TexCoordsV(.NumMappedVerts - 1)
                    Get fileNumber, offset, .TexCoordsV
                    offset = offset + 2 * 4 * CLng(.NumMappedVerts)
                Case &H4160:
                    'Local transformation matrix
                    Get fileNumber, offset, .LocalTransformationMatrix
                    offset = offset + 12 * 4
                Case &H4000:
                    'Object
                    offset = offset - 6
                    doneQ = True
                Case Else:
                    'Unknown chunk
                    If offset - 6 >= count Then
                        offset = offset - 6
                        doneQ = True
                    Else
                        offset = offset + llen - 6
                        doneQ = offset >= fileLength
                    End If
            End Select
        Loop Until doneQ
    End With
    
    If temp_mesh.NumVerts > 0 Then
        If SafeArrayGetDim(MeshesV) <> 0 Then
            mesh_index = UBound(MeshesV) + 1
        Else
            mesh_index = 0
        End If
        ReDim Preserve MeshesV(mesh_index)
        MeshesV(mesh_index) = temp_mesh
        'Debug.Print test_str; " "; temp_mesh.NumVerts; " "; temp_mesh.NumFaces
    End If
End Sub
Sub ReadObject3Ds(ByVal fileNumber As Integer, ByRef offset As Long, ByVal fileLength As Long, ByVal Length As Long, ByRef ModelsV() As Model3Ds)
    Dim count As Long
    Dim id As Integer
    Dim llen As Long
    Dim doneQ As Boolean
    Dim ci As Integer
    Dim model_index As Integer
    
    count = offset + Length - 6
    
    doneQ = False
    
    If SafeArrayGetDim(ModelsV) <> 0 Then
        model_index = UBound(ModelsV) + 1
    Else
        model_index = 0
    End If
    ReDim Preserve ModelsV(model_index)
    With ModelsV(model_index)
        Do
            Get fileNumber, offset, id
            offset = offset + 2
            
            If (offset >= fileLength) Then
                doneQ = True
                Exit Do
            End If
            
            Get fileNumber, offset, llen
            offset = offset + 4
            
            Select Case id
                Case &H4000:
                    'Some object chunk (provably a mesh)
                    ReadMesh3Ds fileNumber, offset, fileLength, llen, .MeshesV
                Case &HAFFF:
                    'Material chunk
                    ReadMaterial3Ds fileNumber, offset, fileLength, llen, .MaterialsV
                Case Else:
                    'Unknown chunk
                    If offset - 6 >= count Then
                        offset = offset - 6
                        doneQ = True
                    Else
                        offset = offset + llen - 6
                        doneQ = offset >= fileLength
                    End If
            End Select
        Loop Until doneQ
    End With
End Sub

Sub Read3DS(ByVal fileNumber As Integer, ByRef offset As Long, ByVal fileLength As Long, ByRef ModelsV() As Model3Ds)
    Dim id As Integer
    Dim llen As Long
    Dim doneQ As Boolean
    Dim ci As Integer
    Dim mesh_index As Integer
    
    doneQ = False
    Do
        Get fileNumber, offset, id
        offset = offset + 2
        
        If (offset >= fileLength) Then
            doneQ = True
            Exit Do
        End If
        
        Get fileNumber, offset, llen
        offset = offset + 4
        
        Select Case id
            Case &HFFFF:
                doneQ = True
            Case &H3D3D:
                'Object chunk
                ReadObject3Ds fileNumber, offset, fileLength, llen, ModelsV
            Case Else:
                'Unknown chunk
                offset = offset + llen - 6
                doneQ = offset >= fileLength
        End Select
    Loop Until doneQ
End Sub

Function ReadPrimaryChunk3DS(ByVal fileNumber As Integer, ByRef offset As Long, ByVal fileLength As Long, ByRef ModelsV() As Model3Ds) As Boolean
    Dim version As Byte
    Dim flag As Integer
    
    Get fileNumber, offset, flag
    offset = offset + 4
    If flag = &H4D4D Then
        offset = 29
        Get fileNumber, offset, version
        offset = offset + 1
        'If version < 3 Then
            'Invalid version
        '    ReadPrimaryChunk3DS = False
        'Else
            offset = 17
            Read3DS fileNumber, offset, fileLength, ModelsV
            ReadPrimaryChunk3DS = True
        'End If
    Else
        ReadPrimaryChunk3DS = False
    End If
End Function

Sub BuildFaceMaterialsList(ByRef Model As Model3Ds)
    'Build the list of material indices for every face
    
    Dim num_meshes As Integer
    Dim num_materials As Integer
    Dim mei As Integer
    Dim mai As Integer
    Dim mfi As Integer
    Dim ci As Integer
    Dim fi As Integer
    Dim foundQ As Boolean
    Dim num_face_mat_groups As Integer
    Dim num_faces As Integer
    
    num_meshes = UBound(Model.MeshesV) + 1
    num_materials = UBound(Model.MaterialsV) + 1
    For mei = 0 To num_meshes - 1
        With Model.MeshesV(mei)
            num_face_mat_groups = UBound(.FaceMaterialsV) + 1
            ReDim .FaceMaterialIndicesV(Model.MeshesV(mei).NumFaces - 1)
            For mfi = 0 To num_face_mat_groups - 1
                mai = 0
                foundQ = False
                Do
                    ci = 0
                    While .FaceMaterialsV(mfi).MaterialName(ci) = _
                            Model.MaterialsV(mai).MaterialName(ci) And _
                            .FaceMaterialsV(mfi).MaterialName(ci) <> 0 And _
                            Model.MaterialsV(mai).MaterialName(ci) <> 0
                            ci = ci + 1
                    Wend
                    
                    foundQ = .FaceMaterialsV(mfi).MaterialName(ci) = _
                            Model.MaterialsV(mai).MaterialName(ci)
                    mai = mai + 1
                Loop Until foundQ Or mai = num_materials
        
                mai = mai - 1
                
                num_faces = .FaceMaterialsV(mfi).NumEntries
                For fi = 0 To num_faces - 1
                    .FaceMaterialIndicesV(.FaceMaterialsV(mfi).facesV(fi)) = mai
                Next fi
            Next mfi
        End With
    Next mei
End Sub

Sub Load3DS(ByVal fileName As String, ByRef ModelsV() As Model3Ds)
    Dim fileNumber As Long
    Dim offset As Long
    Dim fileLength As Long
    Dim num_models As Integer
    Dim mi As Integer
    
    On Error GoTo errorH
    fileNumber = FreeFile
    offset = 1
    fileLength = FileLen(fileName)
    Open fileName For Binary As fileNumber
    While (ReadPrimaryChunk3DS(fileNumber, offset, fileLength, ModelsV))
    Wend
    
    If SafeArrayGetDim(ModelsV) <> 0 Then
        num_models = UBound(ModelsV) + 1
    Else
        num_models = 0
    End If
    
    For mi = 0 To num_models - 1
        BuildFaceMaterialsList ModelsV(mi)
    Next mi
    Exit Sub
errorH:
    'Debug.Print "Can't read 3Ds file: "; fileName
End Sub
'---------------------------------------------------------------------------------------------------------
'---------------------------------------- 3Ds => PModel --------------------------------------------------
'---------------------------------------------------------------------------------------------------------
Private Sub GetVerts(ByRef Mesh As mesh_object_node, ByRef vertsV() As Point3D)
    ReDim vertsV(Mesh.NumVerts - 1)
    
    CopyMemory vertsV(0), Mesh.vertsV(0), CLng(Mesh.NumVerts) * 4 * 3
End Sub
Private Sub GetFaces(ByRef Mesh As mesh_object_node, ByRef facesV() As PPolygon)
    Dim fi As Integer
    
    ReDim facesV(Mesh.NumFaces - 1)
    
    For fi = 0 To Mesh.NumFaces - 1
        With facesV(fi)
            .Tag1 = 0
            .Verts(0) = Mesh.facesV(fi).vertC
            .Verts(1) = Mesh.facesV(fi).vertB
            .Verts(2) = Mesh.facesV(fi).vertA
            .Tag2 = &HCFCEA00
        End With
    Next fi
End Sub
Private Sub GetTexCoords(ByRef Mesh As mesh_object_node, ByRef tex_coordsV() As Point2D)
    If Mesh.NumMappedVerts > 0 Then
        ReDim tex_coordsV(Mesh.NumVerts - 1)
        
        CopyMemory tex_coordsV(0), Mesh.TexCoordsV(0), CLng(Mesh.NumMappedVerts) * 2 * 4
    End If
End Sub
Private Sub GetVColors(ByRef Mesh As mesh_object_node, ByRef MaterialsV() As mat_list_node, ByRef vcolorsV() As color)
    Dim ci As Integer
    Dim fi As Integer
    Dim faces_per_vert() As int_vector
    Dim num_faces As Integer
    Dim temp_r As Long
    Dim temp_g As Long
    Dim temp_b As Long
    
    Dim v_index As Integer
    Dim face_index As Integer
    
    ReDim faces_per_vert(Mesh.NumVerts - 1)
    
    For fi = 0 To Mesh.NumFaces - 1
        v_index = Mesh.facesV(fi).vertA
        face_index = faces_per_vert(v_index).Length
        ReDim Preserve faces_per_vert(v_index).vector(face_index)
        faces_per_vert(v_index).vector(face_index) = fi
        faces_per_vert(v_index).Length = face_index + 1
        
        v_index = Mesh.facesV(fi).vertB
        face_index = faces_per_vert(v_index).Length
        ReDim Preserve faces_per_vert(v_index).vector(face_index)
        faces_per_vert(v_index).vector(face_index) = fi
        faces_per_vert(v_index).Length = face_index + 1
        
        v_index = Mesh.facesV(fi).vertC
        face_index = faces_per_vert(v_index).Length
        ReDim Preserve faces_per_vert(v_index).vector(face_index)
        faces_per_vert(v_index).vector(face_index) = fi
        faces_per_vert(v_index).Length = face_index + 1
    Next fi
    
    ReDim vcolorsV(Mesh.NumVerts - 1)
    
    For ci = 0 To Mesh.NumVerts - 1
        temp_r = 0
        temp_g = 0
        temp_b = 0
        
        For fi = 0 To faces_per_vert(ci).Length - 1
            With MaterialsV(Mesh.FaceMaterialIndicesV(faces_per_vert(ci).vector(fi))).Diffuse
                temp_r = temp_r + .red
                temp_g = temp_g + .green
                temp_b = temp_b + .blue
            End With
        Next fi
        
        If (Not faces_per_vert(ci).Length = 0) Then
            With vcolorsV(ci)
                .r = temp_r / faces_per_vert(ci).Length
                .g = temp_g / faces_per_vert(ci).Length
                .B = temp_b / faces_per_vert(ci).Length
                .a = 255
            End With
        End If
    Next ci
End Sub
Private Sub GetPColors(ByRef Mesh As mesh_object_node, ByRef MaterialsV() As mat_list_node, ByRef pcolorsV() As color)
    Dim ci As Integer
    
    ReDim pcolorsV(Mesh.NumFaces - 1)
    
    For ci = 0 To Mesh.NumFaces - 1
        With MaterialsV(Mesh.FaceMaterialIndicesV(ci)).Diffuse
            pcolorsV(ci).r = pcolorsV(ci).r + .red
            pcolorsV(ci).g = pcolorsV(ci).g + .green
            pcolorsV(ci).B = pcolorsV(ci).B + .blue
        End With
    Next ci
End Sub
Private Sub ConvertMesh3DsToPModel(ByRef Mesh As mesh_object_node, ByRef MaterialsV() As mat_list_node, ByRef Model_out As PModel)
    Dim vertsV() As Point3D
    Dim facesV() As PPolygon
    Dim tex_coordsV() As Point2D
    Dim vcolorsV() As color
    Dim pcolorsV() As color
    
    GetVerts Mesh, vertsV
    GetFaces Mesh, facesV
    GetTexCoords Mesh, tex_coordsV
    GetVColors Mesh, MaterialsV, vcolorsV
    GetPColors Mesh, MaterialsV, pcolorsV
    
    AddGroup Model_out, vertsV, facesV, tex_coordsV, vcolorsV, pcolorsV
End Sub
Private Sub ConvertModel3DsToPModel(ByRef Model As Model3Ds, ByRef Model_out As PModel)
    Dim mi As Integer
    Dim num_meshes As Integer
    
    num_meshes = UBound(Model.MeshesV) + 1
    With Model
        For mi = 0 To num_meshes - 1
            ConvertMesh3DsToPModel .MeshesV(mi), .MaterialsV, Model_out
        Next mi
    End With
End Sub

Sub ConvertModels3DsToPModel(ByRef ModelsV() As Model3Ds, ByRef Model_out As PModel)
    Dim mi As Integer
    Dim num_models As Integer
    
    num_models = UBound(ModelsV) + 1
    For mi = 0 To num_models - 1
        ConvertModel3DsToPModel ModelsV(mi), Model_out
    Next mi
    With Model_out
        .head.off00 = 1
        .head.off04 = 1
        
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
    
    ComputeNormals Model_out
    ComputeBoundingBox Model_out
    ComputeEdges Model_out
End Sub

