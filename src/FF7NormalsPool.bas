Attribute VB_Name = "FF7PModelNormalsPool"
Option Explicit
Sub ReadNormals(ByVal NFile As Integer, ByVal offset As Long, ByRef Normals() As Point3D, ByVal NumNormals As Long)
    If NumNormals > 0 Then
        ReDim Normals(NumNormals - 1)
        Get NFile, offset, Normals
    Else
        ReDim Normals(0)
    End If
End Sub
Sub WriteNormals(ByVal NFile As Integer, ByVal offset As Long, ByRef Normals() As Point3D)
    If UBound(Normals()) > 0 Then _
        Put NFile, offset, Normals
End Sub
Sub MergeNormals(ByRef n1() As Point3D, ByRef n2() As Point3D)
    Dim NumNormalsNI1 As Integer
    Dim NumNormalsNI2 As Integer

    NumNormalsNI1 = UBound(n1) + 1
    NumNormalsNI2 = UBound(n2) + 1
    ReDim Preserve n1(NumNormalsNI1 + NumNormalsNI2 - 1)

    CopyMemory n1(NumNormalsNI1), n2(0), NumNormalsNI2 * 12
End Sub
