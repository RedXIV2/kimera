Attribute VB_Name = "FF7PModelEdgesPool"
Option Explicit
Type PEdge
    Verts(1) As Integer
End Type
Sub ReadEdges(ByVal NFile As Integer, ByVal offset As Long, ByRef Edges() As PEdge, ByVal NumEdges As Long)
    ReDim Edges(NumEdges)
    Get NFile, offset, Edges
End Sub
Sub WriteEdges(ByVal NFile As Integer, ByVal offset As Long, ByRef Edges() As PEdge)
    Put NFile, offset, Edges
End Sub
Sub MergeEdges(ByRef e1() As PEdge, ByRef e2() As PEdge)
    Dim NumEdgesE1 As Integer
    Dim NumEdgesE2 As Integer

    NumEdgesE1 = UBound(e1) + 1
    NumEdgesE2 = UBound(e2) + 1
    ReDim Preserve e1(NumEdgesE1 + NumEdgesE2 - 1)

    CopyMemory e1(NumEdgesE1), e2(0), NumEdgesE2 * 4
End Sub
