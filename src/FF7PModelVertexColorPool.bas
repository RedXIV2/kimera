Attribute VB_Name = "FF7PModelVertexColorPool"
Option Explicit
Sub ReadVColors(ByVal NFile As Integer, ByVal offset As Long, ByRef vcolors() As color, ByVal NumVColors As Long)
    ReDim vcolors(NumVColors - 1)
    Get NFile, offset, vcolors
End Sub
Sub WriteVColors(ByVal NFile As Integer, ByVal offset As Long, ByRef vcolors() As color)
    Put NFile, offset, vcolors
End Sub
Sub MergeVColors(ByRef vc1() As color, ByRef vc2() As color)
    Dim NumVColorsVC1 As Integer
    Dim NumVColorsVC2 As Integer

    NumVColorsVC1 = UBound(vc1) + 1
    NumVColorsVC2 = UBound(vc2) + 1
    ReDim Preserve vc1(NumVColorsVC1 + NumVColorsVC2 - 1)

    CopyMemory vc1(NumVColorsVC1), vc2(0), NumVColorsVC2 * 4
End Sub
Sub CopyVColors(ByRef vcolors() As color, ByRef vcolors_out() As color)
    Dim vi As Integer

    Dim numColors As Integer

    numColors = UBound(vcolors) + 1
    ReDim vcolors_out(numColors)

    CopyMemory vcolors_out(0), vcolors(0), numColors * 4
End Sub
Sub SetVColorsAlphaMAX(ByRef vcolors() As color)
    Dim num_colors As Long
    Dim ci As Long

    num_colors = UBound(vcolors) + 1
    For ci = 0 To num_colors - 1
        vcolors(ci).a = 128
    Next ci
End Sub
