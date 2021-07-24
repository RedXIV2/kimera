Attribute VB_Name = "FF7PModelVerticesPool"
Option Explicit
Sub ReadVerts(ByVal NFile As Integer, ByRef Verts() As Point3D, ByVal NumVerts As Long)
    ReDim Verts(NumVerts - 1)
    Get NFile, &H81, Verts
End Sub
Sub WriteVerts(ByVal NFile As Integer, ByRef Verts() As Point3D)
    Put NFile, &H81, Verts
End Sub
Sub MergeVerts(ByRef v1() As Point3D, ByRef v2() As Point3D)
    Dim NumVertsV1 As Integer
    Dim NumVertsV2 As Integer

    NumVertsV1 = UBound(v1) + 1
    NumVertsV2 = UBound(v2) + 1
    ReDim Preserve v1(NumVertsV1 + NumVertsV2 - 1)

    CopyMemory v1(NumVertsV1), v2(0), NumVertsV2 * 12
End Sub
Function GetVertexProjectedCoords(ByRef Verts() As Point3D, ByVal vi As Integer) As Point3D
    glClear GL_DEPTH_BUFFER_BIT
    GetVertexProjectedCoords = GetProjectedCoords(Verts(vi))
End Function
Function GetVertexProjectedDepth(ByRef Verts() As Point3D, ByVal vi As Integer) As Single
    glClear GL_DEPTH_BUFFER_BIT
    GetVertexProjectedDepth = GetDepthZ(Verts(vi))
End Function
Sub DrawVertT(ByRef Verts() As Point3D, ByVal vi As Integer)
    glBegin GL_POINTS
        With Verts(vi)
            glVertex3f .x, .y, .z
        End With
    glEnd
End Sub
