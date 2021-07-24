Attribute VB_Name = "FF7PModelTextureCoords"
Option Explicit
Sub ReadTexCoords(ByVal NFile As Integer, ByVal offset As Long, ByRef TexCoords() As Point2D, ByVal NumTexCoords As Long)
    Dim tci As Long

    If NumTexCoords > 0 Then
        ReDim TexCoords(NumTexCoords - 1)
        Get NFile, offset, TexCoords
    Else
        ReDim TexCoords(0)
    End If
End Sub
Sub WriteTexCoords(ByVal NFile As Integer, ByVal offset As Long, ByRef TexCoords() As Point2D)
    If UBound(TexCoords()) > 0 Then Put NFile, offset, TexCoords
End Sub
Sub MergeTexCoords(ByRef t1() As Point2D, ByRef t2() As Point2D)
    Dim NumTexT1 As Integer
    Dim NumTexT2 As Integer

    NumTexT1 = UBound(t1) + 1
    NumTexT2 = UBound(t2) + 1
    ReDim Preserve t1(NumTexT1 + NumTexT2 - 1)

    CopyMemory t1(NumTexT1), t2(0), NumTexT2 * 4 * 2
End Sub
