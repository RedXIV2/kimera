Attribute VB_Name = "FF7PModelPolygonsPool"
Option Explicit
Type PPolygon
    Tag1 As Integer
    Verts(2) As Integer
    Normals(2) As Integer
    Edges(2) As Integer
    Tag2 As Long
End Type
Sub ReadPolygons(ByVal NFile As Integer, ByVal offset As Long, ByRef Polygons() As PPolygon, ByVal NumPolygons As Long)
    ReDim Polygons(NumPolygons - 1)
    Get NFile, offset, Polygons
End Sub
Sub WritePolygons(ByVal NFile As Integer, ByVal offset As Long, ByRef Polygons() As PPolygon)
    Put NFile, offset, Polygons
End Sub
Sub MergePolygons(ByRef p1() As PPolygon, ByRef p2() As PPolygon)
    Dim NumPolysP1 As Integer
    Dim NumPolysP2 As Integer
    
    NumPolysP1 = UBound(p1) + 1
    NumPolysP2 = UBound(p2) + 1
    ReDim Preserve p1(NumPolysP1 + NumPolysP2 - 1)
    
    CopyMemory p1(NumPolysP1), p2(0), NumPolysP2 * 24
End Sub
