Attribute VB_Name = "FF7PModelHeader"
Option Explicit
Type PHeader
    off00 As Long
    off04 As Long
    VertexColor As Long
    NumVerts As Long
    NumNormals As Long
    off14 As Long
    NumTexCs As Long
    NumNormInds As Long
    NumEdges As Long
    NumPolys As Long
    off28 As Long
    off2c As Long
    mirex_h As Long
    NumGroups As Long
    mirex_g As Long
    off3c As Long
    unknown(16) As Long
End Type

Sub ReadHeader(ByVal NFile As Integer, ByRef head As PHeader)
    ''Debug.Print "Loading P Header"
    With head
        Get NFile, 1, .off00
        Get NFile, 1 + 4, .off04
        If .off00 <> 1 Or .off04 <> 1 Then
            'Debug.Print "Not a valid P file!!!"
            Exit Sub
        End If
        Get NFile, 1 + 4 * 2, .VertexColor
        ''Debug.Print "   VertexColor=" + Str$(.VertexColor)
        Get NFile, 1 + 4 * 3, .NumVerts
        ''Debug.Print "   NumVerts=" + Str$(.NumVerts)
        Get NFile, 1 + 4 * 4, .NumNormals
        ''Debug.Print "   NumNormals=" + Str$(.NumNormals)
        Get NFile, 1 + 4 * 5, .off14
        Get NFile, 1 + 4 * 6, .NumTexCs
        ''Debug.Print "   NumTexCs=" + Str$(.NumTexCs)
        Get NFile, 1 + 4 * 7, .NumNormInds
        ''Debug.Print "   NumNormInds=" + Str$(.NumNormInds)
        Get NFile, 1 + 4 * 8, .NumEdges
        ''Debug.Print "   NumEdges=" + Str$(.NumEdges)
        Get NFile, 1 + 4 * 9, .NumPolys
        ''Debug.Print "   NumPolys=" + Str$(.NumPolys)
        Get NFile, 1 + 4 * 10, .off28
        Get NFile, 1 + 4 * 11, .off2c
        Get NFile, 1 + 4 * 12, .mirex_h
        ''Debug.Print "   mirex_h=" + Str$(.mirex_h)
        Get NFile, 1 + 4 * 13, .NumGroups
        ''Debug.Print "   NumGroups=" + Str$(.NumGroups)
        Get NFile, 1 + 4 * 14, .mirex_g
        ''Debug.Print "   mirex_g=" + Str$(.mirex_g)
        Get NFile, 1 + 4 * 15, .off3c
        Get NFile, 1 + 4 * 16, .unknown
    End With
''Debug.Print "Done"
End Sub
Sub WriteHeader(ByVal NFile As Integer, ByRef head As PHeader)
    With head
        Put NFile, 1, .off00
        Put NFile, 1 + 4, .off04
        Put NFile, 1 + 4 * 2, .VertexColor
        Put NFile, 1 + 4 * 3, .NumVerts
        Put NFile, 1 + 4 * 4, .NumNormals
        Put NFile, 1 + 4 * 5, .off14
        Put NFile, 1 + 4 * 6, .NumTexCs
        Put NFile, 1 + 4 * 7, .NumNormInds
        Put NFile, 1 + 4 * 8, .NumEdges
        Put NFile, 1 + 4 * 9, .NumPolys
        Put NFile, 1 + 4 * 10, .off28
        Put NFile, 1 + 4 * 11, .off2c
        Put NFile, 1 + 4 * 12, .mirex_h
        Put NFile, 1 + 4 * 13, .NumGroups
        Put NFile, 1 + 4 * 14, .mirex_g
        Put NFile, 1 + 4 * 15, .off3c
        Put NFile, 1 + 4 * 16, .unknown
    End With
End Sub
Sub MergeHeader(ByRef h1 As PHeader, ByRef h2 As PHeader)
    With h1
        .NumVerts = .NumVerts + h2.NumVerts
        .NumNormals = .NumNormals + h2.NumNormals
        .NumTexCs = .NumTexCs + h2.NumTexCs
        .NumNormInds = .NumNormInds + h2.NumNormInds
        .NumEdges = .NumEdges + h2.NumEdges
        .NumPolys = .NumPolys + h2.NumPolys
        .mirex_h = .mirex_h + h2.mirex_h
        .NumGroups = .NumGroups + h2.NumGroups
    End With
End Sub
